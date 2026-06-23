from __future__ import annotations

import streamlit as st

from driver_custom_attributes_app import render_driver_custom_attributes_tab


def main() -> None:
    st.set_page_config(page_title="Driver Custom Attributes", layout="wide")
    st.title("Driver Custom Attributes")
    st.write("Import, edit, and clean Optibus driver custom attribute timelines from CSV or Excel templates.")
    render_driver_custom_attributes_tab()


if __name__ == "__main__":
    main()
