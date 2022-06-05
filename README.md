# README

This repository contains utility code for administrative tasks for the
Scouting BuLa "mova".

## Setup

In order to run the scripts, the correct python packages need to be installed.
It is advised to create a new virtual environment (e.g. via conda `conda create --name myenv`) and then
install required packages with

```shell
pip install -r requirements.txt
```

## Export personal events

Export everything into a single files xlsx file:

```shell
python scripts/export-events-per-person.py \
    "./data/Staff Planning - master for DB support.xlsx" \
    "./data/App-Users mit device-ID.xlsx" \
    "./outputs/all_events_merged.xlsx"
```

Export into individual xlsx files (named with the `profile_id`, e.g. `./outputs/events_per_person/{profile_id}.xlsx`):

```shell
python scripts/export-events-per-person.py \
    --split-exports \
    "./data/Staff Planning - master for DB support.xlsx" \
    "./data/App-Users mit device-ID.xlsx" \
    "./outputs/events_per_person"
```