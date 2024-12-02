import streamlit as st
import filter
import straight
import balloon

# Create tabs
tabs = st.tabs(['Filter', 'MTA', 'Balloon'])

# Display content based on the selected tab
with tabs[0]:
    filter.app1_ui()  # Run the content of app1.py

with tabs[1]:
    straight.main()  # Run the content of app2.py

with tabs[2]:
    balloon.main()  # Run the content of app3.py