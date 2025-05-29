import streamlit as st

if 'medical_visited' not in st.session_state:
    st.session_state.medical_visited = False

st.title("Rham Finance Team Application")

if not st.session_state.medical_visited:
    if st.button("Discovery Medical Aid"):
        st.session_state.medical_visited = True
        st.success("Redirecting... Please click the link below if not redirected automatically.")
        st.markdown("[Go to Medical app](http://localhost:8502)")
else:
    st.success("Welcome back! Click the link below to go to Medical app:")
    st.markdown("[Go to Medical app](http://localhost:8502)")
