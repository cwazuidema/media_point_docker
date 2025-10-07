import pandas as pd
import numpy as np
import os

"""
Local processing entrypoint compatible with the previous Lambda interface.
Reads "/data/Bron.xlsx" and writes "/data/Modified_Bron.xlsx".
"""

# Local file paths inside the container (mounted from host)
local_file_path = "/data/Bron.xlsx"
modified_file_path = "/data/Modified_Bron.xlsx"


def lambda_handler(event, context):
    try:
        # Process the local file and write the output locally
        process_excel_file(local_file_path, modified_file_path)
        return {"statusCode": 200, "body": "File processed successfully."}
    except Exception as e:
        return {"statusCode": 500, "body": f"Failed to process file: {e}"}


def process_excel_file(input_file_path, output_file_path):
    df = pd.read_excel(input_file_path).fillna(0)
    df.columns = df.columns.str.lower()
    df["land"] = df["land"].replace("Nederland", "")
    df["email"] = df["email"].str.lower()
    df["postcode"] = df["postcode"].astype(str)

    df.insert(0, "card", 2)
    df.insert(0, "fam", np.nan)
    df.insert(0, "naam compleet", np.nan)
    df.insert(0, "straat compleet", np.nan)
    df.insert(0, "plaats compleet", np.nan)
    df.insert(0, "name", np.nan)
    df.insert(0, "MailChimp", False)
    df.insert(0, "fam_number", 0)
    df.insert(0, "digitaal 2p+ family", False)
    df.insert(0, "digitaal 2p+", False)
    df.insert(0, "digitaal 1p", False)
    df.insert(0, "digitaal", False)
    df.insert(0, "fysiek 2p+ brieven", False)
    df.insert(0, "fysiek 2p+", False)
    df.insert(0, "fysiek 1p", False)
    df.insert(0, "fysiek", False)

    df["name"] = (
        df["tussenvoegsel"]
        .replace(0, "")
        .apply(lambda x: x.strip() + " " if x.strip() != "" else "")
        + df["naam"]
    )
    df["naam compleet"] = (
        df["voornaam"]
        + df["tussenvoegsel"]
        .replace(0, "")
        .apply(lambda x: " " + x if x.strip() != "" else "")
        + " "
        + df["naam"]
    )
    df["straat compleet"] = (
        df["straat"]
        + " "
        + df["huisnummer"].astype(str)
        + df["toevoeging"]
        .replace(0, "")
        .replace("", "")
        .apply(lambda x: " " + x if x.strip() != "" else "")
    )
    df["plaats compleet"] = df["postcode"] + "  " + df["plaats"]
    df["postcode huisnummer toevoeging"] = (
        df["postcode"]
        + " "
        + df["huisnummer"].astype(str)
        + df["toevoeging"]
        .replace(0, "")
        .replace("", "")
        .apply(lambda x: " " + x if x.strip() != "" else "")
    )
    df["fam"] = df["naam compleet"] + " " + df["abonneenummer"].astype(str)
    df["geboortedatum"] = pd.to_datetime(df["geboortedatum"]).dt.strftime("%d-%m-%Y")
    df["vanaf"] = pd.to_datetime(df["vanaf"]).dt.strftime("%d-%m-%Y")

    # Correct logic for fysiek classification based on postcode huisnummer toevoeging
    address_unique = ~df.duplicated(
        subset=["postcode huisnummer toevoeging"], keep=False
    )
    address_duplicate = df.duplicated(
        subset=["postcode huisnummer toevoeging"], keep=False
    )

    df["fysiek"] = df["pas fysiek"] == "Ja"
    df["fysiek 1p"] = address_unique & (df["pas fysiek"] == "Ja")
    df["fysiek 2p+"] = address_duplicate & (df["pas fysiek"] == "Ja")

    filtered_df = df[df["fysiek 2p+"]]
    unique_contracts = filtered_df["contractnummer"].unique()

    for contract in unique_contracts:
        subset = filtered_df[filtered_df["contractnummer"] == contract]
        if (subset["toorts"] == 1).any():
            first_toorts = subset[subset["toorts"] == 1].index[0]
            df.loc[first_toorts, "fysiek 2p+ brieven"] = True
        else:
            if (subset["toorts"] == 0).all():
                oldest = subset[
                    subset["geboortedatum"] == subset["geboortedatum"].min()
                ].index[0]
                df.loc[oldest, "fysiek 2p+ brieven"] = True

    df["digitaal"] = df["pas digitaal"] == "Ja"
    df["digitaal 1p"] = df["digitaal"] & ~df["email"].duplicated(keep=False)
    df["digitaal 2p+"] = df["digitaal"] & df["email"].duplicated(keep=False)

    digital_2p_plus_df = df[df["digitaal 2p+"]]
    unique_emails = digital_2p_plus_df["email"].unique()

    for email in unique_emails:
        subset = digital_2p_plus_df[digital_2p_plus_df["email"] == email]
        if (subset["toorts"] == 1).any():
            main_row_index = subset[subset["toorts"] == 1].index[0]
        else:
            main_row_index = subset["geboortedatum"].idxmin()
        df.loc[main_row_index, "MailChimp"] = True
        joining_rows = subset.index.difference([main_row_index])
        df.loc[joining_rows, "MailChimp"] = False

    df.loc[df["digitaal 1p"], "MailChimp"] = True

    digital_2p_plus_family_df = df[df["digitaal 2p+"]].copy()
    unique_emails_digital_family = digital_2p_plus_family_df["email"].unique()

    for email in unique_emails_digital_family:
        subset = digital_2p_plus_family_df[digital_2p_plus_family_df["email"] == email]
        if (subset["toorts"] == 1).any():
            main_row_index = subset[subset["toorts"] == 1].index[0]
            main_email = subset.loc[main_row_index, "email"]
            matching_emails_indices = subset[
                (subset["email"] == main_email) & (subset.index != main_row_index)
            ].index
            df.loc[matching_emails_indices, "digitaal 2p+ family"] = True
            family_subset = subset[subset["email"] != main_email]
            if not family_subset.empty:
                oldest_family_member_index = family_subset["geboortedatum"].idxmin()
                df.loc[oldest_family_member_index, "digitaal 2p+ family"] = False
                remaining_family_indices = family_subset.index.difference(
                    [oldest_family_member_index]
                )
                df.loc[remaining_family_indices, "digitaal 2p+ family"] = True
        else:
            main_row_index = subset["geboortedatum"].idxmin()
            main_email = subset.loc[main_row_index, "email"]
            matching_emails_indices = subset[
                (subset["email"] == main_email) & (subset.index != main_row_index)
            ].index
            df.loc[matching_emails_indices, "digitaal 2p+ family"] = True

    digitaal_2p_plus_family_df = df[df["digitaal 2p+ family"]].copy()
    unique_contracts = digitaal_2p_plus_family_df["email"].unique()

    for contract in unique_contracts:
        subset = digitaal_2p_plus_family_df[
            digitaal_2p_plus_family_df["email"] == contract
        ]
        subset["geboortedatum"] = pd.to_datetime(
            subset["geboortedatum"], format="%d-%m-%Y"
        )
        subset_sorted = subset.sort_values(by="geboortedatum")

        for i, idx in enumerate(subset_sorted.index):
            df.at[idx, "fam_number"] = i + 1

    fysiek_columns = {
        "Naam compleet": "naam compleet",
        "Geboortedatum": "geboortedatum",
        "Straat compleet": "straat compleet",
        "Plaats compleet": "plaats compleet",
        "Land": "land",
        "Vanaf": "vanaf",
        "Abonneenummer": "abonneenummer",
        "toorts": "toorts",
    }

    fysiek_data = df[df["fysiek"]].copy()
    fysiek_df = fysiek_data[list(fysiek_columns.values())]
    fysiek_df.columns = list(fysiek_columns.keys())

    fysiek_1p_data = df[df["fysiek 1p"]].copy()
    fysiek_1p_df = fysiek_1p_data[list(fysiek_columns.values())]
    fysiek_1p_df.columns = list(fysiek_columns.keys())

    fysiek_2p_plus_data = df[df["fysiek 2p+"]].copy()
    fysiek_2p_plus_df = fysiek_2p_plus_data[list(fysiek_columns.values())]
    fysiek_2p_plus_df.columns = list(fysiek_columns.keys())

    fysiek_2p_plus_brieven_data = df[df["fysiek 2p+ brieven"]].copy()
    fysiek_2p_plus_brieven_df = fysiek_2p_plus_brieven_data[
        list(fysiek_columns.values())
    ]
    fysiek_2p_plus_brieven_df.columns = list(fysiek_columns.keys())

    digitaal_columns = {
        "contractnummer": "contractnummer",
        "cardNumber": "abonneenummer",
        "name": "naam compleet",
        "birthday": "geboortedatum",
        "email": "email",
        "dynamicField": "vanaf",
        "card": "card",
    }

    digitaal_data = df[df["digitaal"]].copy()
    digitaal_df = digitaal_data[list(digitaal_columns.values())]
    digitaal_df.columns = list(digitaal_columns.keys())

    digitaal_1p_data = df[df["digitaal 1p"]].copy()
    digitaal_1p_df = digitaal_1p_data[list(digitaal_columns.values())]
    digitaal_1p_df.columns = list(digitaal_columns.keys())

    digitaal_2p_plus_data = df[
        (df["digitaal 2p+"] == True) & (df["digitaal 2p+ family"] == False)
    ].copy()
    digitaal_2p_plus_df = digitaal_2p_plus_data[list(digitaal_columns.values())]
    digitaal_2p_plus_df.columns = list(digitaal_columns.keys())

    digitaal_2p_plus_df["fam1"] = ""
    digitaal_2p_plus_df["fam2"] = ""
    digitaal_2p_plus_df["fam3"] = ""
    digitaal_2p_plus_df["fam4"] = ""
    digitaal_2p_plus_df.loc[:, ["fam1", "fam2", "fam3", "fam4"]] = digitaal_2p_plus_df[
        ["fam1", "fam2", "fam3", "fam4"]
    ].astype(str)

    family_true_df = df[df["digitaal 2p+ family"] == True]

    for idx, row in family_true_df.iterrows():
        fam_number = row["fam_number"]
        if fam_number in [1, 2, 3, 4]:
            fam_column = f"fam{int(fam_number)}"
            matching_email_rows = digitaal_2p_plus_df["email"] == row["email"]
            digitaal_2p_plus_df.loc[matching_email_rows, fam_column] = row["fam"]

    mail_chimp_columns = {
        "Email Address": "email",
        "First": "voornaam",
        "Name": "name",
        "Lidmaatschapsnummer": "abonneenummer",
    }

    mail_chimp_data = df[df["MailChimp"]].copy()
    mail_chimp_df = mail_chimp_data[list(mail_chimp_columns.values())]
    mail_chimp_df.columns = list(mail_chimp_columns.keys())

    # Remove duplicate emails in MailChimp
    mail_chimp_df = mail_chimp_df.drop_duplicates(subset=["Email Address"])

    writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")
    df.to_excel(writer, index=False, sheet_name="Main")

    fysiek_df.to_excel(writer, index=False, sheet_name="Fysiek")
    fysiek_1p_df.to_excel(writer, index=False, sheet_name="Fysiek 1p")
    fysiek_2p_plus_df.to_excel(writer, index=False, sheet_name="Fysiek 2p+")
    fysiek_2p_plus_brieven_df.to_excel(
        writer, index=False, sheet_name="Fysiek 2p+ brieven"
    )

    digitaal_df.drop(columns=["contractnummer"], inplace=True)
    digitaal_df.to_excel(writer, index=False, sheet_name="Digitaal")

    digitaal_1p_df.drop(columns=["contractnummer"], inplace=True)
    digitaal_1p_df.to_excel(writer, index=False, sheet_name="Digitaal 1p")

    digitaal_2p_plus_df.drop(columns=["contractnummer"], inplace=True)
    digitaal_2p_plus_df.to_excel(writer, index=False, sheet_name="Digitaal 2p+")

    mail_chimp_df.to_excel(writer, index=False, sheet_name="MailChimp")

    workbook = writer.book
    main_worksheet = writer.sheets["Main"]

    fysiek_worksheet = writer.sheets["Fysiek"]
    fysiek_1p_worksheet = writer.sheets["Fysiek 1p"]
    fysiek_2p_plus_worksheet = writer.sheets["Fysiek 2p+"]
    fysiek_2p_plus_brieven_worksheet = writer.sheets["Fysiek 2p+ brieven"]

    digitaal_worksheet = writer.sheets["Digitaal"]
    digitaal_1p_worksheet = writer.sheets["Digitaal 1p"]
    digitaal_2p_plus_worksheet = writer.sheets["Digitaal 2p+"]
    mail_chimp_worksheet = writer.sheets["MailChimp"]

    for i, width in enumerate(get_col_widths(mail_chimp_df)):
        mail_chimp_worksheet.set_column(i, i, width + 1)

    for i, width in enumerate(get_col_widths(df)):
        main_worksheet.set_column(i, i, width + 1)

    for i, width in enumerate(get_col_widths(fysiek_df)):
        fysiek_worksheet.set_column(i, i, width + 1)

    for i, width in enumerate(get_col_widths(fysiek_1p_df)):
        fysiek_1p_worksheet.set_column(i, i, width + 1)

    for i, width in enumerate(get_col_widths(fysiek_2p_plus_df)):
        fysiek_2p_plus_worksheet.set_column(i, i, width + 1)

    for i, width in enumerate(get_col_widths(fysiek_2p_plus_brieven_df)):
        fysiek_2p_plus_brieven_worksheet.set_column(i, i, width + 1)

    for i, width in enumerate(get_col_widths(digitaal_df)):
        digitaal_worksheet.set_column(i, i, width + 1)

    for i, width in enumerate(get_col_widths(digitaal_1p_df)):
        digitaal_1p_worksheet.set_column(i, i, width + 1)

    for i, width in enumerate(get_col_widths(digitaal_2p_plus_df)):
        digitaal_2p_plus_worksheet.set_column(i, i, width + 1)

    writer.close()

    print("File processed and saved successfully with additional sheets.")


def get_col_widths(dataframe):
    return [
        max([len(str(s)) for s in dataframe[col].values] + [len(col)])
        for col in dataframe.columns
    ]
