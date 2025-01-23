import os
from dotenv import load_dotenv
from telegram import Update, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
import pandas as pd
import datetime

# Load environment variables
load_dotenv()
TOKEN = os.getenv("TELEGRAM_TOKEN")

# Initialize datetime for XML generation
dt = datetime.datetime.now()

# Temporary storage for user interactions
user_data = {}

# Function to generate XML
def generate_xml(file_input, output_dir):
    # Read the Excel file
    df = pd.read_excel(file_input)

    # Generate a timestamp-based file name
    output_file_name = f"APY_EXIT_{dt.strftime('%Y%m%d_%H%M%S')}.xml"
    result_output = os.path.join(output_dir, output_file_name)

    with open(result_output, 'w') as filex:
        # XML Header
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

        for _, row in df.iterrows():  # Iterate through all rows
            row_dict = {}  # Initialize a new dictionary for each row
            for key, value in row.to_dict().items():
                if pd.api.types.is_numeric_dtype(df[key]):  # Check if the column is numeric
                    if pd.notna(value):  # Ensure the value is not NaN
                        row_dict[key] = str(int(value)) if value == int(value) else str(value)
                    else:
                        row_dict[key] = ""  # Handle NaN values explicitly
                else:
                    row_dict[key] = str(value)
            
            xd = str.split(row['EXIT_DATE'], " ")
            # Write data for each row in the XML
            filex.write(f'''
<req-dtl>
<pran>{row_dict['PRAN_NO']}</pran>
<wdr-due-to>EN</wdr-due-to>
<wdr-type>P</wdr-type>
<share-to-wdr>100</share-to-wdr>
<share-to-annuity>0</share-to-annuity>
<exit-date>{xd[0]}</exit-date>
<reason-of-closure>3</reason-of-closure>
<specify>I am not interested please close my account</specify>
<subs-bank-dtls>
<bank-ifs-flag>Y</bank-ifs-flag>
<account-no>{row_dict['ACC_NO']}</account-no>
<bank-ifs-code>PUNB0SUPGB5</bank-ifs-code>
<bank-micr-code></bank-micr-code>
<bank-name>PUPGB</bank-name>
<bank-branch>{row_dict['SOL_NO']}</bank-branch>
<bank-address>{row_dict['BRANCH']}</bank-address>
<bank-pin>{row_dict['PINCODE']}</bank-pin>
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

# Command Handlers
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Welcome to the APY EXIT Converter Bot! Please upload an Excel file to start."
    )

async def file_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user
    user_id = user.id
    user_data[user_id] = {}

    # Save uploaded file
    file = update.message.document
    file_name = file.file_name
    file_path = os.path.join("uploads", file_name)

    os.makedirs("uploads", exist_ok=True)
    
    # Properly await `get_file()` and then download the file
    file_obj = await file.get_file()  # Await the coroutine to get the file object
    await file_obj.download_to_drive(file_path)  # Download the file

    user_data[user_id]['file_input'] = file_path

    await update.message.reply_text(
        "File received! Generating the XML file based on your input..."
    )

    try:
        # Create output directory
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)

        # Generate XML
        result_path = generate_xml(file_path, output_dir)

        # Send the generated XML back to the user
        with open(result_path, 'rb') as result_file:
            await update.message.reply_document(
                document=InputFile(result_file),
                filename=os.path.basename(result_path),
                caption="XML file generated successfully!"
            )

        # Cleanup temporary files
        os.remove(file_path)
        os.remove(result_path)

    except KeyError as e:
        await update.message.reply_text(f"Missing column in the Excel file: {str(e)}")
    except Exception as e:
        await update.message.reply_text(f"An error occurred: {str(e)}")


# Main Function to Run the Bot
def main():
    # Initialize the Application
    application = Application.builder().token(TOKEN).build()

    # Add command and message handlers
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.Document.ALL, file_handler))

    # Start the bot
    application.run_polling()

if __name__ == "__main__":
    main()
