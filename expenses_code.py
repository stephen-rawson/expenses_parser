import datetime as dt
import os
import re

import email
import extract_msg
import numpy as np
import pandas as pd

def test_string_inclusion(line, strings):
    test = [s in line.lower() for s in strings]
    return all(test)


def get_weekday(date):
    mapping = dict(
        zip(
            [0, 1, 2, 3, 4, 5, 6],
            [
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
                "Friday",
                "Saturday",
                "Sunday",
            ],
        )
    )
    return mapping[date.weekday()]


def setup_df():
    df = pd.DataFrame(
        columns=[
            "Date",
            "Sender",
            "Subject",
            "FX",
            "Price",
            "From",
            "To",
            "TicketNumber",
        ],
        index=np.arange(0, 10000, 1),
    )
    return df


def get_messages(message_dir):
    messages = [
        os.path.join(message_dir, f)
        for f in os.listdir(message_dir)
        if f.endswith(".msg")
    ]
    return messages


def open_message(filepath):
    return extract_msg.Message(filepath)


def identify_sender(msg):
    sender = msg.sender
    if "deliveroo" in sender.lower():
        return "Deliveroo"
    if "uber receipts" in sender.lower():
        return "Uber Travel"
    if "bcdtravel" in sender.lower():
        return "BCD"
    return "Unknown"


def parse_deliveroo(body):
    content = [line for line in body.split("\r\n") if line != ""]
    for index, line in enumerate(content):
        if "total" == line.lower().strip() and len(line) < 10:
            fx = content[index + 1].strip().split(" ")[0]
            price = float(content[index + 1].strip().split(" ")[1])
            break
    return fx, price


def parse_uber_travel(body):
    content = [line for line in body.split("\r\n") if line != ""]
    for index, line in enumerate(content):
        if "switch" == line.lower().strip()[:6]:
            fx = re.search("[a-zA-z]+", content[index + 1])[0]
            price = float(re.search("[0-9]{1,5}.[0-9]{0,2}", content[index + 1])[0])
        if "invite your friends and family" in line.lower():
            departure = content[index - 4].strip().replace("\t", "")
            arrival = content[index - 2].strip().replace("\t", "")
    return fx, price, departure, arrival


def parse_bcd(body):
    fx, price, departure, arrival, ticket = [""] * 5
    content = [line for line in body.split("\r\n") if line not in ["", "\t"]]

    for index, line in enumerate(content):
        if "total amount" in line.lower():
            fx = re.search("[a-zA-z]+", content[index][-10:])[0]
            price = float(
                re.search("[0-9,.]{1,15}", content[index][-20:])[0].replace(",", "")
            )
        if test_string_inclusion(line, ["flight", "vendor", "status"]):
            departure, arrival = content[index + 1].split("\t")[1].split("-")
        if test_string_inclusion(line, ["electronic", "ticket", "number"]):
            ticket = content[index + 1].replace("\t", " ").split(" ")[0].strip()
        if test_string_inclusion(line, ["airline", "record",  "locator"]):
            record = line[-15:].split(" ")[-1].replace("\t", "").strip()
            break
    return fx, price, departure, arrival, f"{ticket} - {record}"


def parse_body(sender, body):
    fx, price, departure, arrival, ticket, skipped = [np.nan] * 5 + [False]
    if sender == "Deliveroo":
        fx, price = parse_deliveroo(body)
    elif sender == "Uber Travel":
        fx, price, departure, arrival = parse_uber_travel(body)
    elif sender == "BCD":
        fx, price, departure, arrival, ticket = parse_bcd(body)
    else:
        skipped = True
    return fx, price, departure, arrival, ticket, skipped


def classify_purpose(sender):
    if "BCD" in sender:
        return "* Airfare"
    elif "Deliveroo" in sender:
        return "* Meals Self"
    elif "Uber Travel" in sender:
        return "Taxi"
    else:
        return np.nan

    
def parse_message_list(message_list):
    df = setup_df()
    index = 0
    skipped_count = 0
    for message in message_list:
        msg = open_message(message)
        date = pd.to_datetime(msg.date)
        sender = identify_sender(msg)
        subject = msg.subject

        fx, price, departure, arrival, ticket, skipped = parse_body(sender, msg.body)
        if skipped:
            skipped_count += 1

        df.iloc[index] = [date, sender, subject, fx, price, departure, arrival, ticket]
        index += 1

    df.dropna(inplace=True, how="all")
    df["Weekday"] = df["Date"].apply(get_weekday)
    df = df[
        [
            "Date",
            "Weekday",
            "Sender",
            "Subject",
            "FX",
            "Price",
            "From",
            "To",
            "TicketNumber",
        ]
    ]

    df["Date"] = df.Date.dt.tz_convert(None)
    df["Price"] = df["Price"].astype(float)
    df["Purpose"] = df["Sender"].apply(classify_purpose)

    return df.fillna("NA"), skipped_count


def parse(fpath):
    messages = get_messages(fpath)
    df, skipped_count = parse_message_list(messages)

    return df, skipped_count

