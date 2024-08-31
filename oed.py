import streamlit as st
import requests
import json
import os

def search_oed(word):
    app_id = "1ef445a3"
    app_key = "963b716e198d3825f6dbbb8103200edc"
    language = "en-gb"
    url = f"https://od-api-sandbox.oxforddictionaries.com:443/api/v2/search/{language}/{word.lower()}"
    
    headers = {
        "app_id": app_id,
        "app_key": app_key,
    }
    st.write(f"App id: {app_id}")
    st.write(f"App key: {app_key}")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        return data
    else:
        return None

st.title("Oxford English Dictionary Search")

search_word = st.text_input("Enter a word to search:")
search_button = st.button("Search")

if search_button and search_word:
    result = search_oed(search_word)
    st.write(result)
    
    if result:
        st.subheader(f"Definition for '{search_word}':")
        definitions = result['results'][0]['lexicalEntries'][0]['entries'][0]['senses']
        for i, sense in enumerate(definitions, 1):
            st.write(f"{i}. {sense['definitions'][0]}")
    else:
        st.error("Word not found or there was an error with the API request.")