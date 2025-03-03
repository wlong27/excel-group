import streamlit as st
import pandas as pd
import io
from utils import process_excel_file


def main():
    st.set_page_config(page_title="Excel Group Processor",
                       page_icon="📊",
                       layout="wide")

    # Custom CSS with loading animation
    st.markdown("""
        <style>
        .main {
            padding: 2rem;
        }
        .stButton>button {
            width: 100%;
            margin-top: 1rem;
        }
        .success-message {
            padding: 1rem;
            background-color: #D4EDDA;
            color: #155724;
            border-radius: 0.25rem;
            margin: 1rem 0;
        }
        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.3; }
            100% { opacity: 1; }
        }
        .loading-spinner {
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 2rem;
            background: #f8f9fa;
            border-radius: 8px;
            margin: 1rem 0;
            animation: pulse 1.5s ease-in-out infinite;
        }
        .processing-step {
            padding: 0.5rem;
            margin: 0.2rem 0;
            border-radius: 4px;
            background: #f8f9fa;
        }
        .processing-step.active {
            background: #e9ecef;
            border-left: 4px solid #007bff;
        }
        </style>
    """, unsafe_allow_html=True)

    # Header section
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("📊 Excel Group Processor")
        st.markdown("""
        This tool helps you organize your Excel data by:
        1. Grouping rows based on a column you choose
        2. Creating separate Excel files for each group
        3. Providing a ZIP file containing all grouped Excel files
        """)

    # Main content with shadow box effect
    st.markdown("""
        <div style='padding: 2rem; border-radius: 10px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); background-color: white;'>
    """, unsafe_allow_html=True)

    # File upload and header selection section
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1️⃣ Upload Your File")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Supported formats: .xlsx, .xls"
        )

    with col2:
        st.subheader("2️⃣ Select Header Row")
        header_row = st.number_input(
            "Header row index (0-based):",
            min_value=0,
            value=0,
            help="Choose which row contains your column headers (0 means first row)"
        )

    # Process uploaded file
    if uploaded_file is not None:
        try:
            # Custom loading spinner for file reading
            with st.spinner("📂 Reading Excel file..."):
                st.markdown("""
                    <div class='loading-spinner'>
                        <h3>📊 Reading Excel File...</h3>
                    </div>
                """, unsafe_allow_html=True)
                df = pd.read_excel(uploaded_file, header=header_row)

            if df.empty:
                st.error("⚠️ The uploaded file is empty.")
                return

            # Display file info
            st.markdown("---")
            st.subheader("3️⃣ File Preview")
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"📄 Total Rows: {len(df)}")
            with col2:
                st.info(f"📊 Total Columns: {len(df.columns)}")

            st.dataframe(
                df.head(),
                use_container_width=True,
                height=200
            )

            # Column selection
            st.markdown("---")
            st.subheader("4️⃣ Choose Grouping Column")
            columns = df.columns.tolist()
            selected_column = st.selectbox(
                "Select the column to group by:",
                options=columns,
                key="column_selector",
                help="Separate Excel files will be created for each unique value in this column"
            )

            # Show group preview
            if selected_column:
                unique_values = df[selected_column].nunique()
                st.info(f"🔍 This will create {unique_values} separate Excel files based on unique values in '{selected_column}'")

            # Process button
            if st.button("🚀 Process Files", type="primary"):
                # Initialize progress container
                progress_container = st.empty()

                # Processing steps with animated indicators
                with st.spinner("Processing your files..."):
                    try:
                        # Step 1: Data Preparation
                        progress_container.markdown("""
                            <div class='processing-step active'>
                                📊 Preparing data for grouping...
                            </div>
                        """, unsafe_allow_html=True)

                        # Step 2: Creating Files
                        progress_container.markdown("""
                            <div class='processing-step active'>
                                📊 Preparing data for grouping...
                            </div>
                            <div class='processing-step active'>
                                📑 Creating separate Excel files...
                            </div>
                        """, unsafe_allow_html=True)

                        # Process the files and get zip bytes
                        zip_bytes = process_excel_file(df, selected_column)

                        # Step 3: Finalizing
                        progress_container.markdown("""
                            <div class='processing-step active'>
                                📊 Preparing data for grouping...
                            </div>
                            <div class='processing-step active'>
                                📑 Creating separate Excel files...
                            </div>
                            <div class='processing-step active'>
                                🗜️ Creating ZIP archive...
                            </div>
                        """, unsafe_allow_html=True)

                        # Clear progress container
                        progress_container.empty()

                        # Success message and download button
                        st.markdown("""
                            <div class='success-message'>
                                ✅ Files processed successfully! Click below to download the ZIP archive.
                            </div>
                        """, unsafe_allow_html=True)

                        st.download_button(
                            label="📥 Download ZIP Archive",
                            data=zip_bytes,
                            file_name="grouped_excel_files.zip",
                            mime="application/zip",
                            key="download_button"
                        )
                    except Exception as e:
                        st.error(f"❌ Error processing file: {str(e)}")

        except Exception as e:
            st.error(f"❌ An error occurred while reading the file: {str(e)}")
            st.error(
                "Please make sure the file is a valid Excel file and try again."
            )

    # Close the shadow box
    st.markdown("</div>", unsafe_allow_html=True)

    # Footer
    st.markdown("---")
    st.markdown("""
        <div style='text-align: center; color: #666;'>
            Made with ❤️ by Team 114
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()