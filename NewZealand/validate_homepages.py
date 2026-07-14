from __future__ import annotations

from pathlib import Path

import pandas as pd
import requests


WORKBOOK = Path("NZ_Bioinformatics_Mentors_Ranked_Corrected_v2.xlsx")
SHEET = "Full Corrected List"


def find_url_column(df: pd.DataFrame) -> str:
    return max(
        df.columns,
        key=lambda col: df[col].astype(str).str.startswith(("http://", "https://")).sum(),
    )


def check_url(url: str) -> tuple[str, int | None, str]:
    headers = {"User-Agent": "Mozilla/5.0 homepage validator"}
    try:
        response = requests.head(url, allow_redirects=True, timeout=8, headers=headers)
        if response.status_code in {403, 405} or response.status_code >= 500:
            response = requests.get(
                url,
                allow_redirects=True,
                timeout=10,
                headers=headers,
                stream=True,
            )
        if response.status_code < 400:
            return "OK", response.status_code, response.url
        if response.status_code == 403:
            return "PROTECTED", response.status_code, response.url
        return "BAD", response.status_code, response.url
    except requests.RequestException as exc:
        return "BAD", None, exc.__class__.__name__


def main() -> None:
    df = pd.read_excel(WORKBOOK, sheet_name=SHEET).dropna(how="all")
    url_col = find_url_column(df)
    urls = (
        df[url_col]
        .dropna()
        .astype(str)
        .loc[lambda series: series.str.startswith(("http://", "https://"))]
        .drop_duplicates()
    )

    print(f"Workbook: {WORKBOOK}")
    print(f"Sheet: {SHEET}")
    print(f"URL column: {url_col}")
    print(f"URLs: {len(urls)}")

    ok_count = 0
    protected_count = 0
    for url in urls:
        status, code, final_url = check_url(url)
        ok_count += int(status == "OK")
        protected_count += int(status == "PROTECTED")
        code_text = "ERR" if code is None else str(code)
        print(f"{status}\t{code_text}\t{url}\t{final_url}")

    print(f"Summary: {ok_count}/{len(urls)} reachable, {protected_count} protected")


if __name__ == "__main__":
    main()
