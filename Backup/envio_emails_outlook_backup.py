
import win32com.client
import os
import sys
import logging
import datetime
import pandas as pd
from bs4 import BeautifulSoup

# Constants
SHEET_NAMES = ['Emails', 'PT', 'EN', 'ES']

# Log configuration
timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
log_file = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), f'Relatório_{timestamp}.log')
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def read_excel_data(excel_file):
    """
    Read data from the specified Excel file and process it using pandas.

    Arguments:
        excel_file (str): Path to the Excel file.

    Returns:
        DataFrame: A DataFrame containing email data.
    """
    try:
        data = pd.read_excel(excel_file, sheet_name=SHEET_NAMES, skiprows=1)
    except:
        error_message = "Por favor verifique se o ficheiro 'Envio_Emails.xlsx' se encontra na pasta e se contém as folhas necessárias ('Emails', 'PT', 'EN', 'ES')."
        logging.error(error_message)
        print(error_message)
        os.startfile(log_file)
        sys.exit(1)
    
    emails_df = clean_dataframe(data['Emails'].drop(columns=['Unnamed: 0']))
    pt_df = clean_dataframe(data['PT'].drop(columns=['Unnamed: 0', 'Unnamed: 2']))
    en_df = clean_dataframe(data['EN'].drop(columns=['Unnamed: 0', 'Unnamed: 2']))
    es_df = clean_dataframe(data['ES'].drop(columns=['Unnamed: 0', 'Unnamed: 2']))

    validate_emails_sheet(emails_df)
    validate_language_sheets(pt_df, en_df, es_df)
    validate_attachments(emails_df, attachments_folder)
    
    return process_data(emails_df, pt_df, en_df, es_df)

def clean_dataframe(df):
    """
    Clean the DataFrame by replacing NaN values with empty strings.

    Arguments:
        df (DataFrame): The DataFrame to be cleaned.

    Returns:
        DataFrame: The cleaned DataFrame.
    """
    return df.fillna('')

def validate_emails_sheet(emails_df):
    """
    Validate if the required cells in the Emails sheet contain empty values.

    Arguments:
        emails_df (DataFrame): The Emails DataFrame.

    Raises:
        ValueError: If any of the required cells are missing.
    """
    required_columns = ['Nome Completo (obrigatório)', 'Email (obrigatório)', 'Idioma (obrigatório)']
    empty_cells = emails_df[required_columns].eq('').any(axis=1)
    if empty_cells.any():
        empty_indices = empty_cells[empty_cells].index + 3
        error_message = f"Células vazias encontradas nas colunas 'Nome Completo (obrigatório)', 'Email (obrigatório)' e/ou 'Idioma (obrigatório)' da folha 'Emails' nas seguintes linhas: {', '.join(map(str, empty_indices))}"
        logging.error(error_message)
        print(error_message)
        os.startfile(log_file)
        sys.exit(1)

def validate_language_sheets(pt_df, en_df, es_df):
    """
    Validate if the required cells in the specified sheets contain empty values.

    Arguments:
        pt_df (DataFrame): The Portuguese DataFrame.
        en_df (DataFrame): The English DataFrame.
        es_df (DataFrame): The Spanish DataFrame.
    """
    sheets = {'PT': pt_df, 'EN': en_df, 'ES': es_df}
    for language, df in sheets.items():
        subject = df.iloc[0, 0]
        message = df.iloc[0, 1]
        if pd.isna(subject) or pd.isna(message) or not subject.strip() or not message.strip():
            error_message = f"Não há assunto e/ou mensagem na folha '{language}'."
            logging.error(error_message)
            print(error_message)
            os.startfile(log_file)
            sys.exit(1)

def validate_attachments(emails_df, attachments_folder):
    """
    Validate if the attachments specified in the DataFrame exist in the attachments folder.

    Arguments:
        emails_df (DataFrame): The DataFrame containing email data.

    Raises:
        FileNotFoundError: If any attachment specified in the DataFrame does not exist.
    """
    for index, row in emails_df.iterrows():
        attachments = row['Anexo']
        extensions = row['Extensão']
        if attachments and extensions:
            attachments_list = [attachment.strip() for attachment in attachments.split(';')]
            extensions_list = [('.' + extension.strip() if not extension.strip().startswith('.') else extension.strip()) for extension in extensions.split(';')]
            for attachment, extension in zip(attachments_list, extensions_list):
                    attachment_path = os.path.join(attachments_folder, str(attachment) + str(extension))
                    if not os.path.exists(attachment_path):
                        error_message = f"O anexo '{str(attachment) + str(extension)}' da linha {index + 3} não foi encontrado na pasta 'Anexos'"
                        logging.error(error_message)
                        print(error_message)
                        os.startfile(log_file)
                        sys.exit(1)

def process_data(emails_df, pt_df, en_df, es_df):
    """
    Process data from the specified DataFrames and return email data.

    Arguments:
        emails_df (DataFrame): The DataFrame containing email data.
        pt_df (DataFrame): The Portuguese DataFrame.
        en_df (DataFrame): The English DataFrame.
        es_df (DataFrame): The Spanish DataFrame.

    Returns:
        DataFrame: A DataFrame containing email data.
    """
    data_list = []

    for index, row in emails_df.iterrows():
        full_name, company, email_address, cc, bcc, language, attachments, extensions = row[
            ['Nome Completo (obrigatório)', 'Empresa (se aplicável)', 'Email (obrigatório)', 'CC', 'BCC', 'Idioma (obrigatório)', 'Anexo', 'Extensão']
            ]

        full_name = '' if pd.isna(full_name) else full_name
        company = '' if pd.isna(company) else company
        email_address = '' if pd.isna(email_address) else email_address
        cc = '' if pd.isna(cc) else cc
        bcc = '' if pd.isna(bcc) else bcc
        language = '' if pd.isna(language) else language
        attachments = '' if pd.isna(attachments) else attachments
        extensions = '' if pd.isna(extensions) else extensions

        first_name = full_name.split()[0] if ' ' in full_name else full_name
        subject, message = get_subject_and_message(pt_df, en_df, es_df, language)
        validate_message_contains_name(language, message)
        if not company:
            subject = subject.replace('[NOME]', full_name)
        else:
            subject = subject.replace('[NOME]', company)
        message = message.replace('[NOME]', first_name)
        data_list.append({
            'Nome Completo': full_name,
            'Empresa': company,
            'Email': email_address,
            'CC': cc,
            'BCC': bcc,
            'Assunto': subject,
            'Mensagem': message,
            'Anexo': attachments,
            'Extensão': extensions
        })
    return pd.DataFrame(data_list)

def get_subject_and_message(pt_df, en_df, es_df, language):
    """
    Get the subject and message based on the specified language.

    Arguments:
        pt_df (DataFrame): The Portuguese DataFrame.
        en_df (DataFrame): The English DataFrame.
        es_df (DataFrame): The Spanish DataFrame.
        language (str): The language identifier ('PT', 'EN', 'ES').

    Returns:
        tuple: A tuple containing the subject and message.
    """
    df = {'PT': pt_df, 'EN': en_df, 'ES': es_df}.get(language)
    if df is None:
        error_message = f"Idioma não reconhecido na coluna 'Idioma': {language}"
        logging.error(error_message)
        print(error_message)
        os.startfile(log_file)
        sys.exit(1)
    return df.iloc[0, 0], df.iloc[0, 1]

def validate_message_contains_name(language, message):
    """
    Validate if '[NOME]' is present in the message for the specified language.

    Arguments:
        language (str): The language identifier.
        message (str): The message content.
    """
    if '[NOME]' not in message:
        error_message = f"'[NOME]' não está presente na mensagem da folha '{language}'."
        logging.error(error_message)
        print(error_message)
        os.startfile(log_file)
        sys.exit(1)

def is_outlook_running():
    """
    Check if Outlook is running.

    Returns:
        boolean: True if Outlook is running, False otherwise.
    """
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
        return True
    except:
        return False

def send_emails(data_df, attachments_folder):
    """
    Send emails in Outlook based on the provided DataFrame.

    Arguments:
        data_df (DataFrame): A DataFrame containing email data.
        attachments_folder (str): Path to the folder containing attachments.
    """
    if not is_outlook_running():
        error_message = "Por favor inicie o Outlook e tente novamente."
        logging.error(error_message)
        print(error_message)
        os.startfile(log_file)
        sys.exit(1)
    
    outlook = win32com.client.Dispatch("Outlook.Application")

    failed_emails = []
    drafts_created = 0
    processed_emails = set()
    
    has_attachments = data_df['Anexo'].notna().any()
    has_empty_attachments = data_df['Anexo'].eq('').any()

    while True:
        if has_attachments and has_empty_attachments:
            user_input = input("Há emails sem anexo no ficheiro Excel. Enviar email apenas para emails que contenham anexos? (Sim/Não): ").strip().lower()
            if user_input in ['sim', 's']:
                data_df = data_df[data_df['Anexo'].ne('')]
                break
            elif user_input in ['não', 'nao', 'n']:
                break
            else:
                print("Resposta inválida. Por favor, responda apenas com 'Sim' ou 'Não'.")
        else:
            break

    for index, row in data_df.iterrows():
        full_name, company, email_address, cc, bcc, subject, message, attachments, extensions = row
        if email_address in processed_emails:
            error_message = f"Email duplicado, não vai ser enviado email: {email_address}"
            logging.warning(error_message)
            print(error_message)
            continue
        try:
            mail = outlook.CreateItem(0)
            mail.To = email_address
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc
            mail.Subject = subject
            mail.Display()
            signature = mail.HTMLBody
            message = message.replace('\n', '<br>')
            formatted_message = f'<div style="font-family: Arial, sans-serif; font-size: 10pt;">{message}</div>'
            formatted_message = formatted_message.replace('[N]', '<b>').replace('[/N]', '</b>')
            formatted_message = formatted_message.replace('[S]', '<u>').replace('[/S]', '</u>')
            formatted_message = formatted_message.replace('[I]', '<i>').replace('[/I]', '</i>')
            mail.HTMLBody = f'{formatted_message}{signature}'
            
            if attachments and extensions:
                attachments_list = [attachment.strip() for attachment in attachments.split(';')]
                extensions_list = [('.' + extension.strip() if not extension.strip().startswith('.') else extension.strip()) for extension in extensions.split(';')]
                for attachment, extension in zip(attachments_list, extensions_list):
                    attachment_path = os.path.join(attachments_folder, str(attachment) + str(extension))
                    mail.Attachments.Add(attachment_path)
            
            mail.Send()
            drafts_created += 1
            processed_emails.add(email_address)
        except Exception as e:
            failed_emails.append((full_name, email_address))
            print(e)
            
    if drafts_created > 0:
        completion_message = f"{drafts_created} emails enviados."
        logging.info(completion_message)
        print(completion_message)
    
    if failed_emails:
        error_message = "Não foi possível enviar email para os destinatários abaixo:\n"
        for full_name, email_address in failed_emails:
            error_message += f" - {full_name}\n"
        logging.error(error_message)
        print(error_message)

if __name__ == "__main__":

    attachments_folder = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), 'Anexos')
    excel_file = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), 'Envio_Emails.xlsx')
    while True:
        user_input = input("O ficheiro Excel foi gravado antes de iniciar o programa? (Sim/Não): ").strip().lower()
        if user_input in ['sim', 's']:
            inp_message = f"O ficheiro Excel foi gravado antes de iniciar o programa? (Sim/Não): {user_input}"
            logging.info(inp_message)
            break
        elif user_input in ['não', 'nao', 'n']:
            error_message = "Por favor grave o ficheiro Excel e reinicie o programa."
            logging.error(error_message)
            print(error_message)
            os.startfile(log_file)
            sys.exit()
        else:
            print("Resposta inválida. Por favor insira 'Sim' ou 'Não'.")
    print("A recolher dados do ficheiro Excel...")
    data_df = read_excel_data(excel_file)
    print("A enviar emails...")
    send_emails(data_df, attachments_folder)

os.startfile(log_file)
