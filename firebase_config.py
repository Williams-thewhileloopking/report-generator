
import firebase_admin
import streamlit as st
from firebase_admin import credentials, firestore


def initialize_firebase():

    if not firebase_admin._apps:

        if "firebase" in st.secrets:
            cred = credentials.Certificate(
                dict(st.secrets["firebase"])
            )
        else:
            cred = credentials.Certificate(
                "serviceAccountKey.json"
            )

        firebase_admin.initialize_app(cred)

    db = firestore.client()

    return db
