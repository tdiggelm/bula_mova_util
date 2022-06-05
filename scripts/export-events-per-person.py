import argparse
import pandas as pd
import os
from tqdm.autonotebook import tqdm
from itertools import chain


def export(input_filename, app_users_filename, output_path, split_exports):
    checks_start_col = 3
    header_rows = slice(2, 8)
    header_cols = slice(4, None)
    body_rows = slice(9, None)
    body_cols = slice(1, None)

    df_main = pd.read_excel(input_filename)
    df_devices = pd.read_excel(app_users_filename)
    if split_exports:
        os.makedirs(output_path, exist_ok=True)

    device_ids = {}
    for _, row in df_devices.iterrows():
        device_ids[row["profil_id"]] = row["deviceIds"]

    event_ids_per_day = []
    header_data = df_main.iloc[header_rows, header_cols].transpose()
    for row in header_data.itertuples(index=False):
        events = [event for event in row if pd.notna(event)]
        assert len(events) == len(set(events))
        event_ids_per_day.append([event for event in row if pd.notna(event)])

    body_data = df_main.iloc[body_rows, body_cols]
    all_records = []
    for row in tqdm(body_data.itertuples(index=False), desc="Exporting data"):
        name, email, profile_id = row[0:3]
        all_event_ids = []
        for event_ids, check in zip(event_ids_per_day, row[checks_start_col:]):
            if pd.notna(check):
                all_event_ids.extend(event_ids)
        records = []
        for event_id in all_event_ids:
            records.append(
                {
                    "Name": name,
                    "E-Mail": email,
                    "Profil-ID": profile_id,
                    "TerminId": event_id,
                    "_deviceId": device_ids[profile_id],
                }
            )
        all_records.append((profile_id, records))

    columns = ["Name", "E-Mail", "Profil-ID", "TerminId", "_deviceId"]
    
    if split_exports:
        for profile_id, records in all_records:
            personal_df = pd.DataFrame.from_records(records, columns=columns)
            output_filename = os.path.join(output_path, f"{profile_id}.xlsx")
            personal_df.to_excel(output_filename, index=False)
    else:
        merged_records = chain.from_iterable(
            records for _, records in all_records
        )
        merged_df = pd.DataFrame.from_records(merged_records, columns=columns)
        merged_df.to_excel(output_path, index=False)


def get_args():
    parser = argparse.ArgumentParser(description="Export events per person")
    parser.add_argument(
        "input_filename", help="The main file containing all personal and events"
    )
    parser.add_argument(
        "input_app_users_filename", help="Contains a table with users and deviceids"
    )
    parser.add_argument(
        "output_path", help="The target output path: a **directory** that is used for outputting individial xlsx exports for each person if flag `--split_exports is set, otherwise an **filename** where the entire export is written to as xlsx"
    )
    parser.add_argument(
        "--split-exports",
        action="store_true",
        help="writes individual outputs for each person if set, otherwise merges all outputs into single file"
    )
    return parser.parse_args()


def main():
    args = get_args()
    export(
        args.input_filename,
        args.input_app_users_filename,
        args.output_path,
        args.split_exports,
    )


if __name__ == "__main__":
    main()
