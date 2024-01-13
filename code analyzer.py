import datetime
import streamlit as st
import pandas as pd
import nltk
from nltk.tokenize import sent_tokenize
from nltk.corpus import stopwords
from langdetect import detect
from nltk.probability import FreqDist
from exchangelib import Credentials, Account, DELEGATE
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer
from sumy.nlp.stemmers import Stemmer
from datetime import timedelta, datetime
import re
import pytz

# Read the CSV file into a pandas DataFrame
dataset_file = 'client_type.csv'  # Replace with the actual file path
df = pd.read_csv(dataset_file)

# Your credentials and server settings here
USERNAME = ''
PASSWORD = ''
SERVER = 'outlook.office365.com'

# Configuration des informations d'identification
credentials = Credentials(username=USERNAME, password=PASSWORD)

# Connexion à votre compte Outlook 365
account = Account(
    primary_smtp_address=USERNAME,
    credentials=credentials,
    autodiscover=True,
    access_type=DELEGATE
)

st.set_page_config(layout="wide")
morocco_timezone = pytz.timezone('Africa/Casablanca')
morocco_now = datetime.now(morocco_timezone)
time_threshold = morocco_now - timedelta(hours=72)

width = 40  # Adjust this value according to your preference
margin_left = 20  # Adjust the left margin as needed
margin_right = 20  # Adjust the right margin as needed


# Define a function to set the width and margins of Streamlit components
def set_component_width_and_margins(width, margin_left, margin_right):
    st.markdown(f'<style>.stSelectbox {{ width: {width}px; margin-right: {margin_right}px; }}</style>', unsafe_allow_html=True)
    st.markdown(f'<style>.stTextInput {{ width: {width}px; margin-left: {margin_left}px; margin-right: {margin_right}px; }}</style>', unsafe_allow_html=True)
    st.markdown(f'<style>.stButton {{ width: {width}px; margin-left: {margin_left}px; margin-right: {margin_right}px; }}</style>', unsafe_allow_html=True)
    st.sidebar.markdown(f'<style>.st-bb {{ width: {width}px; margin-left: {margin_left}px; margin-right: {margin_right}px; }}</style>', unsafe_allow_html=True)

# Use the set_component_width_and_margins function to set the width and margins
set_component_width_and_margins(width, margin_left, margin_right)


# Function to create a styled header
def styled_header(title):
    return st.write(
        f"<h2 style='text-align:center; color:#333;'>{title}</h2>",
        unsafe_allow_html=True,
    )

# Define a list of keywords or phrases that indicate urgency 
urgency_keywords = [
    r'\bbloquant\b',
    r'\bblocked\b',
    r'\bDown\b',
    r'\bService interrompu\b',
    r'\bindisponible\b',
    r'\bprobleme\s+d\'acces\b',
    r'\bPanne\b',
    r'\bemission et reception \b',
    r'\bInternet Down\b',
    r'\bIncident\b',
    r'\bpas opérationnelle\b',
    r'\bhors serviceb\b']

# Combine the keywords into a single regular expression pattern, and make it case-insensitive
urgency_pattern = '|'.join(urgency_keywords)
# Initialize a dictionary to keep track of sent emails by sender
sent_emails_by_sender = {}
sent_emails_by_sender_subject_day = {}

def tokenize_sentences(content):
    return sent_tokenize(content)

def preprocess(sentences):
    stop_words = set(stopwords.words("english"))

    processed_sentences = []
    for sentence in sentences:
        words = nltk.word_tokenize(sentence.lower())
        words = [word for word in words if word.isalnum() and word not in stop_words]
        processed_sentences.append(" ".join(words))

    return processed_sentences

# Identify and remove headers or footers using common patterns
def remove_headers_footers(content):
    lines = content.split('\n')
    cleaned_lines = []
    skip_next_line = False

    for line in lines:
        if skip_next_line:
            skip_next_line = False
            continue

        # Identify common patterns in headers or footers
        patterns_to_skip = ['From:', 'Sent:', 'To:', 'Subject:', 'Cc:', 'Bcc:', '--', 'Regards', 'Sincerely',
                            'De :', 'Envoyé :', 'À :', 'Objet :', 'Cc :', 'Cci :', '--', 'Cordialement', 'Sincèrement']

        if any(pattern in line for pattern in patterns_to_skip):
            skip_next_line = True
            continue

        cleaned_lines.append(line)

    cleaned_content = '\n'.join(cleaned_lines)
    return cleaned_content

def generate_summary(content, num_sentences=3):
    # Detect language and set appropriate stemmer
    if any(ord(c) > 127 for c in content):
        language = 'french'
        stemmer = Stemmer(language)
    else:
        language = 'english'
        stemmer = Stemmer(language)

    cleaned_content = remove_headers_footers(content)  # Remove headers and footers
    paragraphs = cleaned_content.split('\n\n')  # Tokenize content into paragraphs
    summarizer = LsaSummarizer(stemmer)

    # Create a summary object for each paragraph
    summary_paragraphs = []
    for paragraph in paragraphs:
        parser = PlaintextParser.from_string(paragraph, Tokenizer(language))
        summary_obj = summarizer(parser.document, num_sentences)
        summary = " ".join([str(s) for s in summary_obj])
        summary_paragraphs.append(summary)

    return "\n\n".join(summary_paragraphs)

# Create a dictionary from the DataFrame for client names and types
client_dataset = {row['Client']: row['Type Client (Standard/VIP)'] for _, row in df.iterrows()}

# Define client type priorities
client_type_priority = {
    "OPERATEUR": 1, "GR_PREMIUM": 1,
    "GRANDS COMPTES": 2,"Banques": 2, "GR_4H": 2,
    "STANDARD": 3, "FAUCON": 3,
    "Other Client": 4
}

email_data = []
def main():
    st.title("Email Analyser App")

    
    selected_time_range = st.selectbox("  ", ["All", "Last 24 Hours", "Last 48 Hours", "Last 72 Hours"])
    filter_option = st.radio("  ", ["Critiques","GTR","Bloquants","Standards","All"])

    email_data = []
    urgent_emails = []
    standard_emails = []
    gtr_mails = []
    crtl_mails=[]
    emails_display =[]

    for item in account.inbox.all():
        recipients = [recipient.email_address for recipient in item.to_recipients]
        sender = item.sender.email_address
        sent_date = item.datetime_sent
        received_date = item.datetime_received
        email_subject = item.subject
        if received_date.tzinfo is None:
            received_date = morocco_timezone.localize(received_date)

        # Calculate the time threshold based on the selected time range
        if selected_time_range == "Last 24 Hours":
            time_threshold = morocco_now - timedelta(hours=24)
        elif selected_time_range == "Last 48 Hours":
            time_threshold = morocco_now - timedelta(hours=48)
        elif selected_time_range == "Last 72 Hours":
            time_threshold = morocco_now - timedelta(hours=72)
        else:
            time_threshold = None  # Show all emails

        # Extract the day from the sent_date as a string (YYYY-MM-DD)
        sent_day = str(sent_date.date())

        #Construct a unique key for the combination of sender and email_subject
        sender_subject_day_key = f"{sender}_{email_subject}_{sent_day}"

        # Search for client names in email subject (object)
        found_client = None
        for client_name in client_dataset:
            if client_name in item.subject:
                found_client = client_name
                break

        # If not found in subject, search in email content
        if not found_client:
            for client_name in client_dataset:
                if client_name.lower() in item.text_body.lower():
                    found_client = client_name
                    break


        # Assign the appropriate label based on client presence
        if found_client:
            client_type = client_dataset[found_client]
        else:
            client_type = "Other Client"

        sent_by_sender = 0
        if sender_subject_day_key in sent_emails_by_sender_subject_day:
            sent_by_sender = sent_emails_by_sender_subject_day[sender_subject_day_key]
        else:
            sent_emails_by_sender_subject_day[sender_subject_day_key] = 0
            sent_by_sender = 0

        # Increment the count for this sender_subject_day_key
        sent_emails_by_sender_subject_day[sender_subject_day_key] += 1
        sent_by_sender += 1

        email_content = item.text_body
        summary = generate_summary(email_content)

        email_data.append({
            "recipients": recipients,
            "sender": sender,
            "sent_date": sent_date,
            "received_date": received_date,
            "sent_by_sender": sent_by_sender,
            "summary": summary,
            "email_subject": email_subject,
            "found_client": found_client,
            "client_type": client_type,
            "email_content": email_content,
            "priority": client_type_priority.get(client_type, 999)  # Default to a high priority if not found
        })

    # Sort the email data based on priority and date
    sorted_email_data = sorted(email_data, key=lambda x: (x["sent_date"].timestamp()), reverse=True)

    # Display the sorted emails using Streamlit components
    for email in sorted_email_data:
        email_content = email["email_subject"]

        # Check if the email contains any urgency keywords, case-insensitive
        is_urgent = re.search(urgency_pattern, email_content, re.IGNORECASE) is not None

        if is_urgent:
            email["is_urgent"] = True
            urgent_emails.append(email)
        else:
            email["is_urgent"] = False
            standard_emails.append(email)

    for email in sorted_email_data:
        email_content = email["email_subject"]
        is_urgent = "GTR" in email_content
        is_critical = "CRTL" in email_content
        if is_urgent:
            email["is_urgent"] = True
            gtr_mails.append(email)
        elif is_critical:
            email["is_critical"] = True
            crtl_mails.append(email)
        else:
            if email not in standard_emails:
                standard_emails.append(email)
            

    if filter_option == "All":
        emails_to_display = sorted_email_data
    elif filter_option == "Bloquants":
        emails_to_display = urgent_emails
    elif filter_option == "Standards":
       unique_standard_emails = []
       for email in standard_emails:
          if email not in urgent_emails and email not in gtr_mails and email not in crtl_mails:
            unique_standard_emails.append(email)
          unique_standard_emails.sort(key=lambda x: (x["priority"], x["sent_date"].timestamp()))
          emails_to_display = unique_standard_emails
    elif filter_option == "GTR":
        emails_to_display = gtr_mails
    elif filter_option == "Critiques":
        emails_to_display = crtl_mails
    

    if time_threshold:
        filtered_emails = [email for email in emails_to_display if email["received_date"] >= time_threshold]
    else:
        filtered_emails = emails_to_display

    # Display the filtered emails using Streamlit components
    for email in filtered_emails:
        st.write("=" * 38)
        
        st.write("Source:", email["sender"])
        st.write("Date de réception:", email["received_date"])
        st.write("Objet:", email["email_subject"])
        st.write("Nombre de réclamations:", email["sent_by_sender"])
        st.write("Client:", email["found_client"])
        st.write("Type:", email["client_type"])
        st.write("Priorité:", email["priority"])

# Run the Streamlit app
if __name__ == "__main__":
    main()
