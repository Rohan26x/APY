import os
from dotenv import load_dotenv
from telegram import Update, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, filters, CallbackContext
import pandas as pd
import datetime

# Load environment variables
load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# Function to generate XML
def generate_xml(file_input, fn):
    dt = datetime.datetime.now()
    result_file = fn + ".xml"
    result_output = os.path.join("output", result_file)

    df = pd.read_excel(file_input)

    os.makedirs("output", exist_ok=True)

    with open(result_output, 'w') as filex:
        filex.write('''<?xml version="1.0" encoding="utf-8"?>
<file xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="exitWDR_APY.xsd">
            \n''')

        filex.write(f'''
<header>
<req-type>exitwdr</req-type>
<req-date>{dt.strftime("%y-%m-%d")}</req-date>
<no-of-records>{len(df)}</no-of-records>
<reg-no>7005154</reg-no>
<entity-type>NLOO</entity-type>
<back-offc-ref-num>700515418082205</back-offc-ref-num>
<transaction-type>L</transaction-type>
</header>
\n''')

        for _, row in df.iterrows():
            row = {key: str(value) for key, value in row.to_dict().items()}
            filex.write(f'''
<req-dtl>
<pran>{row['PRAN NO']}</pran>
<wdr-due-to>EN</wdr-due-to>
<wdr-type>P</wdr-type>
<share-to-wdr>100</share-to-wdr>
<share-to-annuity>0</share-to-annuity>
<exit-date>{row['EXIT DATE/STATUS']}</exit-date>
<reason-of-closure>1</reason-of-closure>
<subs-bank-dtls>
<bank-ifs-flag>Y</bank-ifs-flag>
<account-no>{row['ACC NO']}</account-no>
<bank-ifs-code>PUNB0SUPGB5</bank-ifs-code>
<bank-micr-code></bank-micr-code>
<bank-name>PUPGB</bank-name>
<bank-branch>{row['SOL No']}</bank-branch>
<bank-address>{row['BRANCH']}</bank-address>
<bank-pin>{row['PINCODE']}</bank-pin>
<active-bank-account>Y</active-bank-account>
</subs-bank-dtls>
<doc-check-list>
<wdr-doc-list></wdr-doc-list>
<sub-poi-list></sub-poi-list>
<sub-poa-list></sub-poa-list>
</doc-check-list>
</req-dtl>
\n''')

        filex.write("</file>")

    return result_output

async def start(update: Update, context: CallbackContext):
    await update.message.reply_text(
        "Welcome! Please upload your Excel file to convert it to XML."
    )

async def handle_file(update: Update, context: CallbackContext):
    user = update.message.from_user
    file = update.message.document

    # Save uploaded file
    file_name = file.file_name
    file_path = os.path.join("uploads", file_name)

    os.makedirs("uploads", exist_ok=True)
    await file.get_file().download(file_path)

    # Extract filename without extension
    fn = os.path.splitext(file_name)[0]

    # Generate XML
    result_path = generate_xml(file_path, fn)

    # Send the XML back to the user
    with open(result_path, 'rb') as result_file:
        await update.message.reply_document(
            document=InputFile(result_file),
            filename=os.path.basename(result_path),
            caption="Here is your XML file!"
        )

    # Cleanup temporary files
    os.remove(file_path)
    os.remove(result_path)

def main():
    application = Application.builder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_file))

    application.run_polling()

if __name__ == "__main__":
    main()
