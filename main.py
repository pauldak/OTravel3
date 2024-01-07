import openpyxl
from openpyxl.styles import Font

import openai

import streamlit as st
from trymap import generate_google_maps_embed


st.set_page_config(layout="wide")

openai.api_key = st.secrets["OPENAI_API_KEY"]

def save_to_excel(text, start_place, end_place):
    import openpyxl
    from io import BytesIO


    workbook = openpyxl.Workbook()

    # Create in memory BytesIO buffer
    buffer = BytesIO()

    # Get active sheet
    sheet = workbook.active

    rows = text.split("\n")
    for i, row in enumerate(rows):
        cols = row.split(",")
        for j, col in enumerate(cols):
            sheet.cell(row=i + 1, column=j + 1).value = col

    # Iterate over columns to find the maximum width of each column

    for col_id, col in enumerate(sheet.columns, start=2):  # Start from the 2nd column
        max_length = 0
        column = [cell for cell in col]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = max_length
        sheet.column_dimensions[col[0].column_letter].width = adjusted_width

    blue_font = Font(color="0000FF")
    link_row = len(rows)
    sheet["A" + str(link_row)].font = blue_font

    # Adjust the 1st column to a fixed width of 8
    sheet.column_dimensions['A'].width = 15


    # Save Excel file on server

    workbook.save(buffer)
    buffer.seek(0)
    file_name = f'Itinerary {start_place} to {end_place}.xlsx'
    st.download_button(
        label="Download Excel Report",
        data=buffer,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    # st.write("Your file sample.xlsx is ready")

    # return response



def generate_itinerary(start_place, end_place, must_see, max_km, budget, num_days, start_date, selected_pois):
    # Validate
    if not start_place or not end_place:
        # or not terms_checkbox.isChecked():
        st.echo("Error", "Invalid input")
        return

    my_pois_list = selected_pois
    st.echo(my_pois_list)
    num_of_columns = "7"

    user_message = "Generate a table with the following: Plan an itinerary for my upcoming trip:\n"
    round_trip = start_place == end_place
    if round_trip:
        user_message += f', I want to have a round trip  by car. start at  {start_place}  and end at {start_place}\n'
    else:
        user_message += f'from {start_place} to {end_place} by car. '
        user_message += f'I do not want to come back to {start_place} at the end of the trip.\n '

    if len(must_see) > 2:
        user_message += f'At some point during the trip, I must see {must_see} . Not necessarily in the same order. '
        user_message += "You may add additional POIs that you think I might like.\n "
    user_message += "Please OMIT any introductory lines or prefix.\n "
    user_message += "I want to get an itinerary that follow the next rules: \n"
    user_message += "- You can choose the itinerary that you think is the best for me. \n"
    user_message += "- I don't want to arrive to the same place twice, unless it is at the last day of a round trip. \n"
    user_message += f'- The trip will start on {str(start_date)}  . \n'
    user_message += f'- The trip is going to last {str(num_days)} days \n'
    user_message += (f'- It is imperative that I do not drive more than  {str(max_km)} kilometers on any given day'
                     f'-This is a MUST! \n')
    user_message += (f'"- Please distribute driving distances and activities evenly across the days of the trip, '
                     f'avoiding excessive driving on any single day.\n ')
    if my_pois_list:
        user_message += f'- My favorites POIs are:  {str(my_pois_list)} .\n '

    user_message += f'- I do not want to visit in {start_place} . \n'
    user_message += ("- I want to visit 3 or 4 sites every day, total time around 5 to 7 hours per day "
                     "(Depending on the average spending time in each site). \n")
    user_message += "- if there are some POIs on the way, I would like to visit them as well.\n "
    user_message += (f'"- Accommodations with a budget not exceeding {str(budget)} dollars per night, '
                     f'I seek comfortable and welcoming hotel stays"that are rated at least 4.5 stars. ')
    user_message += "- Please check the availability of the hotels before you add them to the itinerary.\n "
    user_message += ("- Provide distinct itinerary for each day of the journey. "
                     "The lines of the table are for the days, \n")
    user_message += "(please separate between the days with a" + r'''\n).'''
    user_message += f'The columns (" {str(num_of_columns)} ) are: '
    user_message += "- Day date (call the column 'Day'). \n"
    user_message += "- Driving from and driving to (in the same row, separate them with ' to ') (call the column 'Way')\n "
    user_message += ("- If we stay in same place DON'T add anything, just write the name of the place, "
                     "refrain from including any character or word before or after.\n ")
    user_message += "- Actual Driving distance (call the column 'km').\n "
    user_message += "- What to do in the morning (with average time in each site) (call the column 'morning') "
    user_message += "if the average time is not integer, round it to the nearest integer. "
    user_message += ("if there are more than one thing to do in the morning, separate them with a '|'. "
                     "Refrain from including any additional commas to the sites names. \n")
    user_message += "- What to do in the afternoon (with average time in each site) (call the column 'afternoon') "
    user_message += "if the average time is not integer, round it up to the nearest integer. "
    user_message += ("if there are more than one thing to do in the afternoon, separate them with a '|'. "
                     "Refrain from including any additional commas to the sites names. \n")
    user_message += "- Hotel name (call the column 'Hotel'). \n"
    user_message += "- Budget (call the column 'Budget'). \n "
    user_message += "SEPARATE between columns with a ',' \n"

    user_message += ("I also need that the first line of the table will be: Day, Way, km,"
                     " morning, afternoon, Hotel, Budget \n")
    user_message += ("At the end of the table, please give me the itinerary in Google Maps format"
                     " with Hyper link and with blue color, "
                     "starts with '=HYPERLINK(")

    hyper_str = "https://www.google.com/maps/dir/'"
    user_message = f'{user_message}"{hyper_str}'

    user_message += (" for each city In the Google Maps format, add its country after the city, "
                     "with a '+' between them. "
                     "Pls refrain from including anything to this link, not before and not after ")

    # st.write(user_message)

    # Call API

    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"Ask the model: '{user_message}' Answer:"},
        ],
        max_tokens=1000
    )

    # Extract the generated answer
    itinerary = response['choices'][0]['message']['content'].strip()

    # Process and export itinerary

    st.write(itinerary)
    save_to_excel(itinerary, start_place, end_place)

st.title("Trip Planner")

# Input fields
# Set the width of the input fields
input_width = 400  # Adjust the width as needed

# Apply custom CSS to set the width of the input fields

st.markdown(
    f"""
    <style>
        .stTextInput, .stNumberInput, .stDateInput, .stMultiselect {{ width: {input_width}px; }}
    </style>
    """,
    unsafe_allow_html=True,
)
# Create text input fields
start_place = st.text_input("Start Place", key="start_place")
end_place = st.text_input("End Place", key="end_place")

must_see = st.text_input("Must See")

# Use beta_expander to create a container for the number input
# Create a sidebar for additional settings
st.sidebar.header("Settings")
max_km = st.sidebar.number_input("Max Km/Day", min_value=150, max_value=300, step=10)
budget = st.sidebar.number_input("Budget Per Night", min_value=150, max_value=1000, step=10)
num_days = st.sidebar.number_input("Number of Days", min_value=1, max_value=10, step=1)
start_date = st.sidebar.date_input("Start Date")

poi_options = ["Museums", "Parks & Gardens", "Architecture", "Art Galleries", "Local festivals",
               "Zoos & Aquariums", "Wineries", "Science Centers", "Local Markets"]

selected_pois = st.sidebar.multiselect("Preferred POIs", poi_options)

terms_checkbox = st.checkbox("I agree to the terms and conditions")

if st.button("Enter Data"):
    with st.spinner("Please wait..."):
        # generate a map from start_place to end_place

        google_maps_embed = generate_google_maps_embed(start_place, end_place)
        st.markdown(google_maps_embed, unsafe_allow_html=True)

        st.write("Your data is being processed. This may take a few moments...")

        # Call your generate_itinerary function with the collected data
        generate_itinerary(start_place, end_place, must_see, max_km, budget, num_days, start_date, selected_pois)
        # st.write("Your itinerary.xlsx is ready in your Downloads directory")
