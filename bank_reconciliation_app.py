import re
from decimal import Decimal, localcontext
from pathlib import Path
from zipfile import BadZipFile

import pandas as pd
import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException

ACCOUNT = "계좌번호"
LEDGER_AMOUNT = "장부가액"
BANK_AMOUNT = "회신금액"
LEDGER_COLUMNS = [ACCOUNT, "은행명", "예금종류", LEDGER_AMOUNT]
BANK_COLUMNS = [ACCOUNT, "은행명", BANK_AMOUNT, "담보설정여부", "비고"]
AMOUNT_PATTERN = re.compile(r"[+-]?(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\.[0-9]+)?")
STATUS_COLORS = {
    "Match": "color: #14532d; background-color: #dcfce7",
    "Discrepancy": "color: #9a3412; background-color: #ffedd5",
    "Missing": "color: #991b1b; background-color: #fee2e2",
}


def read_source(upload, label, columns, amount_column):
    excel = Path(upload.name).suffix.lower() == ".xlsx"
    upload.seek(0)
    if excel:
        data = pd.read_excel(upload, dtype=object, keep_default_na=False, engine="openpyxl")
    else:
        try:
            data = pd.read_csv(upload, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        except UnicodeDecodeError:
            upload.seek(0)
            data = pd.read_csv(upload, dtype=str, keep_default_na=False, encoding="cp949")

    data.columns = data.columns.str.strip()
    missing = [column for column in columns if column not in data.columns]
    if missing:
        raise ValueError(f"{label} 필수 열 누락 ({', '.join(missing)})")
    if not isinstance(data.index, pd.RangeIndex) or not data.columns.is_unique:
        raise ValueError(f"{label}의 열 제목이나 각 행의 열 개수를 확인해 주세요.")
    data = data[columns].copy()

    if excel and not data[ACCOUNT].map(lambda value: isinstance(value, str)).all():
        raise ValueError(f"{label}의 Excel 계좌번호는 원본 번호를 텍스트로 저장해 주세요.")
    data[ACCOUNT] = data[ACCOUNT].astype(str).str.strip()
    if data[ACCOUNT].eq("").any():
        raise ValueError(f"{label}에 계좌번호가 빈 행이 있습니다.")
    duplicate_rows = data[ACCOUNT].duplicated(keep=False)
    if duplicate_rows.any():
        raise ValueError(f"{label}에 중복 계좌번호가 포함된 행이 {duplicate_rows.sum():,}개 있습니다.")

    amounts = data[amount_column].map(
        lambda value: format(Decimal(str(value)), "f") if isinstance(value, float) else str(value)
    ).str.strip()
    invalid = ~amounts.str.fullmatch(AMOUNT_PATTERN)
    if invalid.any():
        raise ValueError(f"{label}의 {amount_column} 열에 빈 값이나 잘못된 금액이 {invalid.sum():,}개 있습니다.")
    data[amount_column] = amounts.str.replace(",", "", regex=False).map(Decimal)
    text_columns = [column for column in columns if column not in (ACCOUNT, amount_column)]
    data[text_columns] = data[text_columns].astype(str)
    return data


def reconcile(ledger, bank):
    result = ledger.merge(
        bank, on=ACCOUNT, how="outer", suffixes=("_장부", "_회신"),
        indicator=True, validate="one_to_one",
    )
    both = result["_merge"].eq("both")
    same_amount = result[LEDGER_AMOUNT].eq(result[BANK_AMOUNT])
    result["Status"] = "Missing"
    result.loc[both & same_amount, "Status"] = "Match"
    result.loc[both & ~same_amount, "Status"] = "Discrepancy"
    result["누락 유형"] = result.pop("_merge").map({
        "left_only": "은행 미회신",
        "right_only": "회사 장부 미등재",
        "both": "",
    })
    result["차액"] = pd.Series(None, index=result.index, dtype=object)
    for row, book, reply in result.loc[both, [LEDGER_AMOUNT, BANK_AMOUNT]].itertuples():
        with localcontext() as context:
            lowest_place = min(book.as_tuple().exponent, reply.as_tuple().exponent)
            context.prec = max(context.prec, max(book.adjusted(), reply.adjusted()) - lowest_place + 2)
            result.at[row, "차액"] = book - reply

    return result[[
        ACCOUNT, "Status", "누락 유형", "은행명_장부", "은행명_회신", "예금종류",
        LEDGER_AMOUNT, BANK_AMOUNT, "차액", "담보설정여부", "비고",
    ]]


def exception_csv(exceptions):
    export = exceptions.copy()
    text_columns = [ACCOUNT, "은행명_장부", "은행명_회신", "예금종류", "담보설정여부", "비고"]
    for column in text_columns:
        # Guard uploaded text that Excel may interpret as a formula.
        export[column] = export[column].map(
            lambda value: "'" + value
            if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@", "＝", "＋", "－", "＠"))
            else value
        )
    for column in [LEDGER_AMOUNT, BANK_AMOUNT, "차액"]:
        export[column] = export[column].map(lambda value: format(value, "f"), na_action="ignore")
    return export.to_csv(index=False).encode("utf-8-sig")


def main():
    st.set_page_config(page_title="은행 잔액 대사", layout="wide")
    st.title("은행 잔액 대사 대시보드")
    st.caption("계좌번호를 기준으로 회사 예금장부와 은행 잔액증명서 회신을 비교합니다.")

    ledger_area, bank_area = st.columns(2)
    ledger_file = ledger_area.file_uploader("회사 예금장부", type=["csv", "xlsx"])
    ledger_area.caption("필수 열 · " + ", ".join(LEDGER_COLUMNS))
    bank_file = bank_area.file_uploader("은행 회신 내역", type=["csv", "xlsx"])
    bank_area.caption("필수 열 · " + ", ".join(BANK_COLUMNS))
    st.caption("Excel은 첫 번째 시트를 읽습니다. 계좌번호는 텍스트로 저장하고 두 파일의 하이픈 표기를 통일해 주세요.")
    if ledger_file is None or bank_file is None:
        st.info("두 파일을 모두 올리면 자동으로 대사합니다.")
        return

    try:
        ledger = read_source(ledger_file, "회사 장부", LEDGER_COLUMNS, LEDGER_AMOUNT)
        bank = read_source(bank_file, "은행 회신", BANK_COLUMNS, BANK_AMOUNT)
    except (UnicodeDecodeError, pd.errors.ParserError, pd.errors.EmptyDataError, BadZipFile, InvalidFileException):
        st.error("파일을 읽을 수 없습니다. 첫 행에 열 제목이 있는 CSV 또는 XLSX 파일인지 확인해 주세요.")
        return
    except ValueError as error:
        st.error(str(error))
        return

    result = reconcile(ledger, bank)
    counts = result["Status"].value_counts()
    total, match, discrepancy, missing = st.columns(4)
    total.metric("전체 계좌 수", f"{len(result):,}", border=True)
    match.metric("일치 건수", f"{counts.get('Match', 0):,}", border=True)
    discrepancy.metric("미일치 건수", f"{counts.get('Discrepancy', 0):,}", border=True)
    missing.metric(
        "미회신·장부 미등재 건수", f"{counts.get('Missing', 0):,}", border=True,
        help="장부에만 있는 은행 미회신과 회신에만 있는 회사 장부 미등재를 합산합니다.",
    )

    st.subheader("계좌별 대사 결과")
    st.caption("차액은 장부가액에서 회신금액을 뺀 값입니다. 한쪽 금액이 없으면 차액은 공란입니다.")
    display = result.copy()
    for column in [LEDGER_AMOUNT, BANK_AMOUNT, "차액"]:
        display[column] = display[column].map(lambda value: format(value, ",f"), na_action="ignore")
    styled = display.style.map(STATUS_COLORS.get, subset=["Status"])
    st.dataframe(styled, hide_index=True, width="stretch")

    exceptions = result.loc[result["Status"].ne("Match")]
    st.download_button(
        "미일치·누락 건 CSV 다운로드", exception_csv(exceptions),
        file_name="bank_reconciliation_exceptions.csv", mime="text/csv",
        disabled=exceptions.empty, on_click="ignore",
    )


if __name__ == "__main__":
    main()
