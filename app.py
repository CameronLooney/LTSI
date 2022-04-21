import os
import streamlit as st
import numpy as np
from PIL import  Image

# Custom imports
from multipage import MultiPage
from pages import main, feedback , issue_form  # import your pages here

# Create an instance of the app
app = MultiPage()

# Title of the main page
display = Image.open('logo.png')

st.image(display, width = 650)
# st.title("Data Storyteller Application")



# Add all your application here
app.add_page("Generate Open Order File", main.app)
app.add_page("Consolidate Feedback", feedback.app)
app.add_page("Issue/Bug Form", issue_form.app)


# The main app
app.run()