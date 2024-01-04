import streamlit as st
from geopy.geocoders import Nominatim


# Function to generate coordinates for a city using geopy

def get_coordinates(city):

    geolocator = Nominatim(user_agent="my-app", timeout=10)  # Increase timeout to 10 seconds
    location = geolocator.geocode(city)
    if location:
        return location.latitude, location.longitude
    else:
        st.error(f"Unable to find coordinates for {city}.")
        return None

# Function to generate pb_parameter based on coordinates
def generate_pb_parameter(coordinates):
    return f"{coordinates[0]}%2C{coordinates[1]}"


# Function to generate the initial part of pb parameter
def generate_initial_pb_part(coordinates_city1):
    return f"{coordinates_city1[0]}%2C{coordinates_city1[1]}!1m28!1m12!1m3!1d5674974.008013442!2d-0.510297724675082!3d46.01152358614516!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!4m13!3e0"


# Function to generate Google Maps embed HTML
def generate_google_maps_embed(city1, city2):
    # Get coordinates for the cities
    coordinates_city1 = get_coordinates(city1)
    coordinates_city2 = get_coordinates(city2)

    if coordinates_city1 and coordinates_city2:
        # Generate the initial part of pb parameter
        pb_initial_part = generate_initial_pb_part(coordinates_city1)

        # Generate pb_parameters
        pb_parameter_city1 = generate_pb_parameter(coordinates_city1)
        pb_parameter_city2 = generate_pb_parameter(coordinates_city2)

        # Generate the embed code
        embed_code = f'<iframe src="https://www.google.com/maps/embed?pb={pb_initial_part}!4m5!1s{pb_parameter_city1}!2s{city1}!3m2!1d{coordinates_city1[0]}!2d{coordinates_city1[1]}!4m5!1s{pb_parameter_city2}!2s{city2}!3m2!1d{coordinates_city2[0]}!2d{coordinates_city2[1]}!5e0!3m2!1sen!2sil!4v1703490655455!5m2!1sen!2sil" width="600" height="450" style="border:0;" allowfullscreen="" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe>'
        return embed_code
