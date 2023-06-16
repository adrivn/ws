from pathlib import Path
import json
import re


def load_json_config(file):
    """
    Loads a JSON file containing external config data, like cell addresses and their corresponding output labels.
    """
    with open(file, "r", encoding="utf8") as f:
        return json.load(f)


def get_missing_values_by_id(
    df, column: str, con, lookup_source: str, column_id_source: str
):
    # Step 1: Split the column into separate rows for each ID
    df = (
        df[column]
        .fillna(0)
        # .astype(str)
        # .str.replace("[()]", "", regex=True)
        # .str.replace("|", ",")
        .str.split(",")
        .explode()
    )
    # Step 2: Load the DataFrame into DuckDB
    con.register("df", df.reset_index())
    # Step 3: Join the DataFrame with the table containing the corresponding values for each ID
    # Step 4: Group by the original index and sum the corresponding values
    query = f"""SELECT df.index, 
    count(*) as count_urs, 
    sum(v.ppa) as sum_ppa,
    sum(v.lsev_dec19) as sum_lsev,
    FROM df
    LEFT JOIN {lookup_source} v ON df.{column} = v.{column_id_source}
    GROUP BY 1
    """
    result = con.execute(query).df()
    result.to_clipboard()
    # Step 5: Return a Series with the sum of corresponding values for each original index
    return result.set_index("index")


def find_files_included(directory, include_pattern):
    p = Path(directory)
    for subdir in p.iterdir():
        if subdir.is_dir() and re.search(include_pattern, str(subdir)):
            for file in subdir.glob("**/[!~$]*.xlsx"):
                if file.is_file():
                    yield file
