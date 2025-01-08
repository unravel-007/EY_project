import pandas as pd
import copy

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

        # Drop the first row
        self.df = self.df.drop(index=0)

        # Get the current columns as a list
        columns = self.df.columns.tolist()
        print(columns)

        # Deep copy of the columns list
        first_row_columns = copy.deepcopy(columns)

        # Iterate through the column names starting from the second column (index 1)
        unnamed_counter = 1
        countt = 1
        for i in range(1, len(first_row_columns)):
            if 'Unnamed:' in str(first_row_columns[i]):  # Check if the column is unnamed
                original_name = first_row_columns[countt]  # The name of the column to its left
                # Rename the unnamed column with a suffix (.1, .2, etc.)
                first_row_columns[i] = f"{original_name}.{unnamed_counter}"
                unnamed_counter += 1
            else:
                countt = i

        self.df.columns = first_row_columns

        # Apply the rearrange logic to the column names
        first_row_columns = self.rearrange_list(first_row_columns)
        self.df = self.df[first_row_columns]
        return self.df

    @staticmethod
    def rearrange_list(A):
        """Rearranges the list of column names."""
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

      # Get the current columns as a list
      columns = self.df.columns.tolist()
      print(columns)
      # Deep copy of the columns list
      first_row_columns = copy.deepcopy(columns)
      self.df.columns = first_row_columns

      # Apply the rearrange logic to the column names
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
      # Start from index 1 to exclude the first element
      if 'Unnamed:' in str(columns[i]):  # Check if the element contains 'Unnamed:'
            first_unnamed_index = i
            break



    # Iterate through the rest of the list
    for i in range(1, len(columns)):
        if 'Unnamed:' not in columns[i]:  # Check if it's a named element
            # Extract the name before the '.' character
            last_named_element = str(columns[i]).split('.')[0]
            for j in range(first_unnamed_index,first_unnamed_index+4):
              if(j==first_unnamed_index):
                columns[j]=last_named_element
              else:
                columns[j]=(f"{last_named_element}.{unnamed_counter}")
                unnamed_counter=unnamed_counter+1
            first_unnamed_index=first_unnamed_index+4
            unnamed_counter = 1  # Reset the unnamed counter when a named element is encountered


    # Print the updated list
    print(columns)
    self.df.columns = columns
    # Apply the rearrange logic to the column names
    columns = self.rearrange_list(columns)
    print(columns)


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





# Read the Excel file
df = pd.read_excel("testing.xlsx")

# Get the column headers as a list
column_headers = df.columns.tolist()

# Display the column headers
print(column_headers)

def check():
  flag=0;
  unnamed_elements = [col for col in column_headers[1:6] if 'Unnamed:' in str(col)]

 # Check if all are unnamed
  if len(unnamed_elements) == 5:
    print("The first 5 elements (excluding the first) are unnamed.")
    return True
  else:
    print("Not all of the first 5 elements (excluding the first) are unnamed.")
    return False


if(check()):
  type=3;
  
elif any('Unnamed:' in str(item) for item in column_headers):
  print("There is an unnamed element in the list.")
  type=1;
else:
  type=2;




if type==1:
  processor = TableStructure1("testing.xlsx")
elif type==2:
  processor = TableStructure2("testing.xlsx")
elif type==3:
  processor = TableStructure3("testing.xlsx")
processor.load_data()
processor.process_table()
processor.save_data("modified_file_with_rearranged_columns.xlsx")
