import pandas as pd
from datetime import date as dt
import os
import flet as ft

def track_job_application(file_path, company_name, job_title, application_status, contact_person, contact_email_or_phone, location, source, follow_up_date, notes):
    today_date = dt.today()

    data = {
        'Company Name': [company_name],
        'Date': [today_date],
        'Job Title': [job_title],
        'Application Status': [application_status],
        'Contact Person': [contact_person],
        'Contact Email/Phone': [contact_email_or_phone],
        'Location': [location],
        'Source': [source],
        'Follow Up Date': [follow_up_date],
        'Notes': [notes]
    }

    df = pd.DataFrame(data)

    try:
        # Load existing data
        existing_df = pd.read_excel(file_path)
        # Append new data
        df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        # File does not exist, create new one
        pass

    # Write the DataFrame to the Excel file
    df.to_excel(file_path, index=False, engine='openpyxl')

    print("Data has been written to the Excel file.")

def main(page: ft.Page):
    # Define the file path
    file_path = r"C:\Users\vishal\OneDrive\Desktop\Job tracking\job_applications.xlsx"
    
    # Create input fields
    company_name_input = ft.TextField(label="Company Name")
    job_title_input = ft.TextField(label="Job Title")
    application_status_input = ft.TextField(label="Application Status")
    contact_person_input = ft.TextField(label="Contact Person")
    contact_email_or_phone_input = ft.TextField(label="Contact Email/Phone")
    location_input = ft.TextField(label="Location")
    source_input = ft.TextField(label="Source")
    follow_up_date_input = ft.TextField(label="Follow Up Date")
    notes_input = ft.TextField(label="Notes", multiline=True)
    
    def on_submit(e):
        # Collect data from input fields
        track_job_application(
            file_path,
            company_name_input.value,
            job_title_input.value,
            application_status_input.value,
            contact_person_input.value,
            contact_email_or_phone_input.value,
            location_input.value,
            source_input.value,
            follow_up_date_input.value,
            notes_input.value
        )
        # Provide feedback to the user
        page.controls.append(ft.Text("Data has been written to the Excel file."))

        #Clear the page
        company_name_input.value = ""
        job_title_input.value = ""
        application_status_input.value = ""
        contact_person_input.value = ""
        contact_email_or_phone_input.value = ""
        location_input.value = ""
        source_input.value = ""
        follow_up_date_input.value = ""
        notes_input.value = ""


        page.update()

    # Create a submit button
    submit_button = ft.ElevatedButton(text="Submit", on_click=on_submit)

    # Arrange components in the page
    page.add(
        company_name_input,
        job_title_input,
        application_status_input,
        contact_person_input,
        contact_email_or_phone_input,
        location_input,
        source_input,
        follow_up_date_input,
        notes_input,
        submit_button
    )

# Run the Flet app
ft.app(target=main)
