"""Module that reads a csv file containing Outlook emails extracted info"""
import os
import argparse
import pandas as pd
from io import StringIO, BytesIO
from aspose.email.storage.pst import PersonalStorage

class PSTMailReader:
    """This class parses an extracted outlook mails pst file and provides tools to work around it"""
    def __init__(self, file=None, unwanted_file: str = 'data/unwanted.csv') -> None:
        self.file = file
        self.unwanted_file: str = unwanted_file
        self.unwanted_list = self.__load_unwanted_list()
        
        self.df = self.__load_pst()

    def __load_unwanted_list(self) -> list:
        """Loads the unwanted senders list from the unwanted file"""
        if os.path.exists(self.unwanted_file):
            with open(self.unwanted_file, 'r', encoding="utf-8") as f:
                return f.read().split(',')
        else:
            return []

    def __load_pst(self) -> pd.DataFrame:
        """Carga los datos desde un archivo PST y devuelve un DataFrame de pandas"""
        if not self.file:
            raise ValueError("No PST file provided")

        try:
            if isinstance(self.file, BytesIO):
                pst = PersonalStorage.from_stream(self.file)
            else:
                pst = PersonalStorage.from_file(self.file)

            count = pst.store.get_total_items_count()
            folder_info = pst.root_folder.get_sub_folder("Bandeja de entrada")

            data = []

            for j in range(0, count, 50):
                print(j)
                message_info_collection = folder_info.get_contents(j,50)
                for message_info in message_info_collection:
                    mapi = pst.extract_message(message_info)

                    sender = mapi.sender_name
                    subject = mapi.subject
                    date = mapi.delivery_time

                    if sender in self.unwanted_list:
                        continue

                    data.append({
                        'Sender': sender,
                        'Subject': subject,
                        'Date': date
                    })

            data = sorted(data, key=lambda x: x['Date'], reverse=True)

            return pd.DataFrame(data)    

        except Exception as e:
            print(f'Error trying to load the pst file: {e}')
            return pd.DataFrame()

class CSVMailReader:
    """This class parses an extracted outlook mails csv file and provides tools to work around it"""
    def __init__(self, file=None, delimiter: str = ',',
                 unwanted_file: str = 'data/unwanted.csv') -> None:
        self.file = file
        self.delimiter: str = delimiter
        self.unwanted_file: str = unwanted_file
        self.df = self.__load_csv()

        self.unwanted_list = self.__load_unwanted_list()
        self.update_senders()

    def __load_csv(self) -> pd.DataFrame:
        """Loads the data from the file object"""
        if self.file is None:
            raise ValueError("No file provided")
        try:
            if isinstance(self.file, StringIO) or hasattr(self.file, 'read'):
                return pd.read_csv(self.file, delimiter=self.delimiter)
            else:
                raise ValueError("Invalid file type")
        except Exception as e:
            print(f"Error loading CSV file: {e}")
            raise

    def __load_unwanted_list(self) -> list:
        """Loads the unwanted senders list from the unwanted file"""
        if os.path.exists(self.unwanted_file):
            with open(self.unwanted_file, 'r', encoding="utf-8") as f:
                return f.read().split(',')
        else:
            return []

    def save_unwanted_list(self) -> None:
        """Saves the unwanted list to the CSV file"""
        with open(self.unwanted_file, 'w', encoding="utf-8") as f:
            f.write(','.join(self.unwanted_list))
        print(f"Unwanted list saved to {self.unwanted_file}")

    def normalize_senders(self):
        """Normalize senders by removing phrases like 'en teams' in multiple languages"""
        teams_variants = ['en teams', 'in teams', 'auf teams', 'sur teams', 'su teams']

        def normalize_name(sender):
            for variant in teams_variants:
                if sender.lower().endswith(variant):
                    return sender[:-(len(variant) + 1)].strip()
            return sender

        self.df['De: (nombre)'] = self.df['De: (nombre)'].apply(normalize_name)
        self.update_senders()

    def update_senders(self):
        """Updates the senders' names and counts"""
        self.senders_df = self.df['De: (nombre)'].value_counts().reset_index()
        self.senders_df.columns = ['Sender', 'Count']

    def print_senders(self) -> None:
        """Prints the senders and the count of how many times they appear"""
        print("List of senders and their counts:")
        for idx, row in self.senders_df.iterrows():
            print(f"{idx + 1}. {row['Sender']} ({row['Count']} times)")

    def remove_sender(self, sender_name: str) -> None:
        """Removes the sender from the dataframe"""
        self.df = self.df[self.df['De: (nombre)'] != sender_name]
        self.update_senders()
        print(f"Sender '{sender_name}' has been removed.")

    def remove_sender_interactive(self, sender_idx: int) -> None:
        """Removes the sender at the given index with user confirmation"""
        sender_to_remove = self.senders_df.iloc[sender_idx - 1]['Sender']
        confirmation = input(f"Are you sure you want to remove all emails from\
                              '{sender_to_remove}'? (y/n): ").lower()
        if confirmation == 'y':
            self.remove_sender(sender_to_remove)
            add_to_unwanted = input("Would you like to add this name to \
                                    the unwanted list? (y/n): ").lower()
            if add_to_unwanted == 'y':
                if sender_to_remove not in self.unwanted_list:
                    self.unwanted_list.append(sender_to_remove)
                    self.save_unwanted_list()
                else:
                    print(f"{sender_to_remove} is already in the unwanted list.")
        else:
            print("Operation cancelled.")

    def load_unwanted_list(self) -> None:
        """Removes all senders from the unwanted list"""
        for sender in self.unwanted_list:
            if sender in self.senders_df['Sender'].values:
                self.remove_sender(sender)
        print("Unwanted senders removed from the dataset.")

    def export_changes(self, output_file: str = None) -> None:
        """Exports the modified data to a new file"""
        if not output_file:
            base, ext = os.path.splitext(self.file_path)
            output_file = f"{base}_modified{ext}"
            count = 1
            while os.path.exists(output_file):
                output_file = f"{base}_modified({count}){ext}"
                count += 1

        self.df.to_csv(output_file, index=False)
        print(f"File exported as: {output_file}")

def show_menu(mail_reader: CSVMailReader):
    """Shows the interactive menu through the console"""
    while True:
        print("\nMenu:")
        print("1. Show senders")
        print("2. Normalize senders (remove 'in Teams' variants)")
        print("3. Remove a sender")
        print("4. Export changes")
        print("5. Exit")

        choice = input("Choose an option: ")
        if choice == '1':
            mail_reader.print_senders()
        elif choice == '2':
            mail_reader.normalize_senders()
            print("Sender names have been normalized.")
        elif choice == '3':
            remove_choice = input("Would you like to remove a sender from the list (1) or load the unwanted list (2)? ")
            if remove_choice == '1':
                mail_reader.print_senders()
                try:
                    sender_idx = int(input("\nSelect the sender number to remove: "))
                    if 1 <= sender_idx <= len(mail_reader.senders_df):
                        mail_reader.remove_sender_interactive(sender_idx)
                    else:
                        print("Invalid sender number. Please try again.")
                except ValueError:
                    print("Invalid input. Please enter a number.")
            elif remove_choice == '2':
                mail_reader.load_unwanted_list()
            else:
                print("Invalid choice.")
        elif choice == '4':
            export_name = input("Enter a file name to export (leave blank for default): ")
            mail_reader.export_changes(export_name)
        elif choice == '5':
            print("Exiting the program.")
            break
        else:
            print("Invalid option. Please try again.")

if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Process an Outlook mail CSV file.")
    parser.add_argument('file_path', metavar='file', type=str, nargs='?', default=None,
                        help='Path to the CSV file containing the mails.')
    parser.add_argument('--delimiter', dest='delimiter', type=str, default=',',
                        help='Optional: delimiter used in the CSV file (default is comma).')

    args = parser.parse_args()

    if not args.file_path:
        print("Usage: python mail_reader.py <file_path> [--delimiter <delimiter>]")
    else:
        # csv_mail_reader_obj = CSVMailReader(args.file_path, args.delimiter)

        # show_menu(csv_mail_reader_obj)
        pst_mail_reader_obj = PSTMailReader(args.file_path)
