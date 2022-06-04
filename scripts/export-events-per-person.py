import argparse
import pandas as pd
import os
from tqdm.autonotebook import tqdm


def export(input_filename, output_dir):
    checks_start_col = 3
    header_rows = slice(2, 8)
    header_cols = slice(4, None)
    body_rows = slice(9, None)
    body_cols = slice(1, None)

    df = pd.read_excel(input_filename)
    os.makedirs(output_dir, exist_ok=True)

    event_ids_per_day = []
    header_data = df.iloc[header_rows, header_cols].transpose()
    for row in header_data.itertuples(index=False):
        events = [event for event in row if pd.notna(event)]
        assert len(events) == len(set(events))
        event_ids_per_day.append([event for event in row if pd.notna(event)])

    body_data = df.iloc[body_rows, body_cols]
    for row in tqdm(body_data.itertuples(index=False), desc="Exporting files"):
        name, email, profile_id = row[0:3]
        all_event_ids = []
        for event_ids, check in zip(event_ids_per_day, row[checks_start_col:]):
            if pd.notna(check):
                all_event_ids.extend(event_ids)
        records = []
        for event_id in all_event_ids:
            records.append({
                "Name": name,
                "E-Mail": email,
                "Profil-ID": profile_id,
                "TerminId": event_id,
            })
        personal_df = pd.DataFrame.from_records(
            records,
            columns=["Name", "E-Mail", "Profil-ID", "TerminId"]
        )
        output_filename = os.path.join(output_dir, f"{profile_id}.xlsx")
        personal_df.to_excel(output_filename)

        
def get_args():
    parser = argparse.ArgumentParser(description="Export events per person")
    parser.add_argument(
        "input_filename",
        help="The main file containing all personal and events"
    )
    parser.add_argument(
        "output_dir",
        help="The target directory for the indivudual exports"
    )
    return parser.parse_args()


def main():
    args = get_args()
    export(args.input_filename, args.output_dir)


if __name__ == '__main__':
    main()