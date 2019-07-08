# Tender Manager

This project is intended to manage incoming tender files at Samskip.

It can pick up a file with an PS1 batch-script, convert the file into a format which our Azure table storage is comfortable with and both upload the original file in the blob storage as well as the converted version in table storage which is then streamlined through our modern data platform to a PowerBI report!

An email may be sent automatically once the process has finished to let other employees know it is done. The PS1 file is stored on a server and runs every 1 hour to check for a new file.

## Dependancies

This project runs mainly on a bunch of different Excel file handling libraries, Azure connector libraries and also some standard Python libraries.

For a full, detailed list, please see `requirements.txt`.

## How-To:

1. Pull this repo to a folder location.
2. Run `pip install -r requirements.txt`.
3. Move the `tender_trigger.ps1` file to a server of your choice and ensure that you use either `windows` task manager or the bulldog extension for `linux` to create an event trigger for the script.
4. Ensure you have also entered your Azure key credentials into the `key_manager.py` on the machine the script will run from. Once this is done remove the values from the Python file. They are now stored on the computer and not within the code.
5. Ensure you have spoken to tender guys to enable them to have a folder which they will place incoming tenders. Once they havew this in place get your PS1 trigger to watch the folder location there.
6. Sit-back and wait for the tender guys to get on with some work... so you can automate it for them!
