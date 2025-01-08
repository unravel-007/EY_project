!pip install gradio
import gradio as gr
import pandas as pd
import copy
import os

class TableProcessor:
    def __init__(self, filepath):
        self.filepath = filepath
        self.df = None

    def load_data(self):
        self.df = pd.read_excel(self.filepath)
        return self.df

    def save_data(self, output_file):
        self.df.to_excel(output_file, index=False)
        print(f"Modified file saved as {output_file}")

    def process_table(self):
        raise NotImplementedError("Subclasses must implement the process_table method.")


class TableStructure1(TableProcessor):
    def process_table(self):
        self.df = self.df.drop(index=0)
        columns = self.df.columns.tolist()
        first_row_columns = copy.deepcopy(columns)
        unnamed_counter = 1
        countt = 1
        for i in range(1, len(first_row_columns)):
            if 'Unnamed:' in str(first_row_columns[i]):
                original_name = first_row_columns[countt]
                first_row_columns[i] = f"{original_name}.{unnamed_counter}"
                unnamed_counter += 1
            else:
                countt = i
        self.df.columns = first_row_columns
        first_row_columns = self.rearrange_list(first_row_columns)
        self.df = self.df[first_row_columns]
        return self.df

    @staticmethod
    def rearrange_list(A):
        i = 0
        while i < len(A):
            current_item = str(A[i])
            if any(current_item in str(A[j]) for j in range(i + 1, min(i + 4, len(A)))):
                found = False
                for j in range(i + 4, len(A)):
                    if current_item in str(A[j]):
                        item_to_move = A.pop(j)
                        A.insert(i + 4, item_to_move)
                        found = True
                        break
                if not found:
                    print(f"No further match found for '{current_item}'.")
            i += 1
        return A


class TableStructure2(TableProcessor):
    def process_table(self):
        columns = self.df.columns.tolist()
        first_row_columns = copy.deepcopy(columns)
        self.df.columns = first_row_columns
        first_row_columns = self.rearrange_list(first_row_columns)
        self.df = self.df[first_row_columns]
        return self.df

    @staticmethod
    def rearrange_list(A):
        i = 0
        while i < len(A):
            current_item = str(A[i])
            if any(current_item in str(A[j]) for j in range(i + 1, min(i + 4, len(A)))):
                found = False
                for j in range(i + 4, len(A)):
                    if current_item in str(A[j]):
                        item_to_move = A.pop(j)
                        A.insert(i + 4, item_to_move)
                        found = True
                        break
                if not found:
                    print(f"No further match found for '{current_item}'.")
            i += 1
        return A


class TableStructure3(TableProcessor):
    def process_table(self):
        columns = self.df.columns.tolist()
        unnamed_counter = 1
        first_unnamed_index = 0
        for i in range(1, len(columns)):
            if 'Unnamed:' in str(columns[i]):
                first_unnamed_index = i
                break
        for i in range(1, len(columns)):
            if 'Unnamed:' not in columns[i]:
                last_named_element = str(columns[i]).split('.')[0]
                for j in range(first_unnamed_index, first_unnamed_index + 4):
                    if j == first_unnamed_index:
                        columns[j] = last_named_element
                    else:
                        columns[j] = f"{last_named_element}.{unnamed_counter}"
                        unnamed_counter += 1
                first_unnamed_index += 4
                unnamed_counter = 1
        self.df.columns = columns
        columns = self.rearrange_list(columns)
        self.df = self.df[columns]
        return self.df

    @staticmethod
    def rearrange_list(A):
        i = 0
        while i < len(A):
            current_item = str(A[i])
            if any(current_item in str(A[j]) for j in range(i + 1, min(i + 4, len(A)))):
                found = False
                for j in range(i + 4, len(A)):
                    if current_item in str(A[j]):
                        item_to_move = A.pop(j)
                        A.insert(i + 4, item_to_move)
                        found = True
                        break
                if not found:
                    print(f"No further match found for '{current_item}'.")
            i += 1
        return A


def process_file(file):
    file_path = file.name
    df = pd.read_excel(file_path)
    column_headers = df.columns.tolist()

    def check():
        unnamed_elements = [col for col in column_headers[1:6] if 'Unnamed:' in str(col)]
        if len(unnamed_elements) == 5:
            return True
        else:
            return False

    if check():
        processor = TableStructure3(file_path)
    elif any('Unnamed:' in str(item) for item in column_headers):
        processor = TableStructure1(file_path)
    else:
        processor = TableStructure2(file_path)

    processor.load_data()
    processor.process_table()
    output_file = "processed_output.xlsx"
    processor.save_data(output_file)
    return output_file


# Gradio interface
with gr.Blocks() as demo:
    gr.Markdown("## Financial Data Sheet")
    with gr.Row():
        file_input = gr.File(label="Upload Excel File", file_types=[".xlsx"])
        file_output = gr.File(label="Download Processed File")
    submit_button = gr.Button("Process File")

    submit_button.click(process_file, inputs=file_input, outputs=file_output)

demo.launch()
