from dotenv import load_dotenv
import streamlit as st
import os
import google.generativeai as genai
import pdfplumber
import fitz  # PyMuPDF
from io import BytesIO
import tempfile
from docx import Document
import re
import all_functions as func
import shutil
import time
from zipfile import ZipFile


# Set page configuration
st.set_page_config(page_title="Resume Data Extractor")
st.header("Resume Data")

api_key = st.text_input('Enter your API key', value='AIzaSyAl5O_YsF9fMFCc1EwSAhwkJOEF704-LJM')
genai.configure(api_key=api_key)

def image_processing_genai(uploaded_files):
    image_data = None
    image_data = func.input_imagedata(uploaded_files)
    query = "Get A to Z details in the resume to put it in text file"
    input_prompt = """
            You are a resume analyzer. You have to extract the data from the resume image. 
            You will have to answer the questions based on the input resume image.
        """
    resume_text = func.get_gemini_response_image(input_prompt, image_data, query)
    return resume_text


    
# File upload function
def file_upload():
    uploaded_files = st.file_uploader("Upload your files...", type=["pdf", "docx", "jpg", "jpeg", "png"], accept_multiple_files=True)
    if uploaded_files is not None:
        return uploaded_files
    else:
        return None

# File processing function
def process_resume(uploaded_files, filename):
    try:
        if uploaded_files.type == "application/pdf":
            resume_text = func.extract_text_from_pdf(uploaded_files)
        elif uploaded_files.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            resume_text = func.extract_text_from_docx(uploaded_files)
        elif uploaded_files.type in ["image/jpeg", "image/png"]:
            resume_text = image_processing_genai(uploaded_files)
            
        else:
            st.error("Unsupported file format. Please upload a PDF, DOCX, JPEG, or PNG file.")
            return
        # Get name response
        name = func.get_name_response(resume_text)
        
        education_overall1 = func.get_education_details_overall(resume_text)
        # Get degree details response
        degree_details = func.get_degree_details_response(education_overall1)
        # Get education details response
        education_details = func.get_education_details_response(education_overall1)
        # Get education years in descending order response
        education_years_descending = func.get_education_years_response(education_overall1)

        
        work_experience = func.get_work_experience_response(resume_text)
        summary = func.get_summary_response(resume_text)     
        # Get certifications response
        certifications = func.get_certifications_response(resume_text)
        # Get technical skills response
        technical_skills = func.get_technical_skills_response2(resume_text)

        
        
        # Fill in the Word document template for all sections
        template_path = 'Templates/agilisium_template.docx'
        output_path = f'agilisium_resume_internal_template/{filename}_resume.docx'
        func.fill_invitation(template_path, output_path, name, summary, certifications)

        # Fill in the degree details inside the table
        func.fill_table_degree_details(output_path, output_path, degree_details)
        # Fill in the institute details inside the table
        func.fill_table_institute_details(output_path, output_path, education_details)
        # Fill in the education years inside the table
        func.fill_table_education_years(output_path, output_path, education_years_descending)
        # Fill in the skill set inside the existing table
        func.fill_table_skill_set(output_path, output_path, technical_skills)
        func.replace_organization_count(output_path, output_path,work_experience)
        
        updated_doc = func.replace_hyphens_with_bullet_points(output_path)
        updated_doc.save(output_path)
        func.remove_characters_from_docx2(output_path)
        updated_doc1 = func.replace_symbol_with_dash(output_path)
        updated_doc1.save(output_path)
        time.sleep(2)
        func.func.process_docx11(output_path)
        func.delete_rows_with_any_empty_cells(output_path)
        st.success(f"{filename}-document processed successfully.")
        
        


    except FileNotFoundError as e:
        st.error(str(e))

   


def process_resume_2(uploaded_files, filename):
    try:
        if uploaded_files.type == "application/pdf":
            resume_text = func.extract_text_from_pdf(uploaded_files)
        elif uploaded_files.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            resume_text = func.extract_text_from_docx(uploaded_files)
        elif uploaded_files.type in ["image/jpeg", "image/png"]:
            resume_text = image_processing_genai(uploaded_files)
        else:
            st.error("Unsupported file format. Please upload a PDF, DOCX, JPEG, or PNG file.")
            return

        
        summary2 = func.get_summary_response(resume_text)
        project_experience = func.relevant_project_experience(resume_text)
        skills1 = func.get_technical_skills_response2(resume_text)
        certifications1 = func.get_certifications_response(resume_text)

        template_path = 'Templates/Client sample format.docx'
        output_path = f'agilisium_resume_client_format/{filename}_resume.docx'
        summar = "DASdSDAas"
        func.fill_invitation2(template_path, output_path,summary2,skills1,project_experience,certifications1,summar)
        education_overall1 = func.get_education_details_overall(resume_text)
        degree_details1 = func.get_degree_details_response(education_overall1)
        education_details1 = func.get_education_details_response(education_overall1)
        education_years_descending1 = func.get_education_years_response(education_overall1)
        func.fill_table_degree_details(output_path, output_path, degree_details1)
        func.fill_table_institute_details(output_path, output_path, education_details1)
        func.fill_table_education_years(output_path, output_path, education_years_descending1)
        updated_doc = func.replace_hyphens_with_bullet_points(output_path)
        updated_doc.save(output_path)

        time.sleep(2)
        func.process_docx11(output_path)
        func.delete_rows_with_any_empty_cells(output_path)
        func.remove_characters_from_docx2(output_path)
        st.success(f"{filename}-document processed successfully.")
        
    except FileNotFoundError as e:
        st.error(str(e))
        



        
def process_resume_3(uploaded_files,filename):
    temp_pdf_path = None
    try:
        if uploaded_files.type == "application/pdf":
            resume_text = func.extract_text_from_pdf(uploaded_files)
        elif uploaded_files.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            resume_text = func.extract_text_from_docx(uploaded_files)
        elif uploaded_files.type in ["image/jpeg", "image/png"]:
            resume_text = image_processing_genai(uploaded_files)
        else:
            st.error("Unsupported file format. Please upload a PDF, DOCX, JPEG, or PNG file.")
            return
        
        # Get name response
        name1 = func.get_name_response(resume_text)
        # Get summary response
        summary3 = func.get_summary_response(resume_text)
        # Get technical skills response
        technical_skills3 = func.get_technical_skills_response2(resume_text)
        
        template_path = 'Templates/Client sample format-2.docx'
        output_path = f'agilisium_resume_client_format_2/{filename}_resume.docx'
        func.fill_invitation3(template_path, output_path, name1, summary3)
        func.fill_table_skill_set(output_path, output_path, technical_skills3)
        education_overall1 = func.get_education_details_overall(resume_text)
        degree_details1 = func.get_degree_details_response(education_overall1)
        education_details1 = func.get_education_details_response(education_overall1)
        education_years_descending1 = func.get_education_years_response(education_overall1)
        
        func.fill_table_degree_details(output_path, output_path, degree_details1)
        func.fill_table_institute_details(output_path, output_path, education_details1)
        func.fill_table_education_years(output_path, output_path, education_years_descending1)
        
        work_experience2 = func.get_work_experience_response2(resume_text)
        func.replace_organization_count2(output_path, output_path,work_experience2)
        
        updated_doc = func.replace_hyphens_with_bullet_points(output_path)
        updated_doc.save(output_path)
        func.remove_characters_from_docx2(output_path)
        updated_doc1 = func.replace_symbol_with_dash(output_path)
        updated_doc1.save(output_path)
        time.sleep(2)
        func.func.process_docx11(output_path)
        func.delete_rows_with_any_empty_cells(output_path)

        # Task from pdf to image
        filename = func.save_document(uploaded_files)
        filename = func.convert_to_pdf_if_docx(filename)
        out_path1 = func.pdf_to_image(filename)
        func.extract_and_save_passport_photo(out_path1)
        image_path = 'passport_photo.png'
        func.replace_placeholder_with_image(output_path,image_path)
        func.remove_pdf_and_docx_files_in_script_directory()
        
        st.success(f"{filename}-document processed successfully.")
        
    except FileNotFoundError as e:
        st.error(str(e))      
# Function to process and save DOCX file
def process_and_save(uploaded_files, process_func, folder_path):
    if uploaded_files:
        os.makedirs(folder_path, exist_ok=True)  # Create a folder to store the documents
        for uploaded_file in uploaded_files:
            filename = os.path.splitext(uploaded_file.name)[0]  # Remove the file extension
            process_func(uploaded_file, filename)
        # Once processing is done, zip and offer the folder for download
        zipped_bytes_io = func.zip_folder_to_bytesio(folder_path)
        st.download_button(
            label="Download Zip Folder",
            data=zipped_bytes_io,
            file_name=f"{folder_path}.zip",
            mime="application/zip"
        )

# Button functions
def internal_template_button(uploaded_files):
    process_and_save(uploaded_files, process_resume, "agilisium_resume_internal_template")

def client_template_button(uploaded_files):
    process_and_save(uploaded_files, process_resume_2, "agilisium_resume_client_format")

def client_template_with_photo_button(uploaded_files):
    process_and_save(uploaded_files, process_resume_3, "agilisium_resume_client_format_2")
    


# File upload
uploaded_files = file_upload()  # Implement file_upload to handle file upload
directories = ["agilisium_resume_internal_template", "agilisium_resume_client_format", "agilisium_resume_client_format_2"]
 # Implement func.delete_directories to clean up before starting

st.sidebar.subheader("Choose Template:")
if st.sidebar.button("Internal Template") and uploaded_files is not None:
    func.delete_directories(directories)
    func.remove_pdf_and_docx_files_in_script_directory()
    internal_template_button(uploaded_files)

if st.sidebar.button("Client Template") and uploaded_files is not None:
    func.delete_directories(directories) 
    func.remove_pdf_and_docx_files_in_script_directory()
    client_template_button(uploaded_files)

if st.sidebar.button("Client Template with Photo") and uploaded_files is not None:
    func.delete_directories(directories)
    func.remove_pdf_and_docx_files_in_script_directory()
    client_template_with_photo_button(uploaded_files)
