"""
Script to create a summary excel file from experiment excel files.
"""

import os
from typing import List, Union
import pandas as pd
import click


def fill_column(summary: List[Union[str, int, float]], sheet: pd.DataFrame, header: str) -> List[Union[str, int, float]]:
    """
    Extracts data from the specified column in the given dataframe and adds it to the summary_data list.

    Args:
        summary_data: A list of summary data that will be appended with data from the given column.
        dataframe: The dataframe containing the column to extract data from.
        column_name: The name of the column in the dataframe to extract data from.

    Returns:
        The updated summary_data list with the data from the specified column added to the end.
    """
    if not sheet[header].empty:
        if sheet[header].isnull().values.any():
            summary.append('empty')
        else:
            summary.append(str(sheet[header][0]))
    else:
        summary.append('empty')
    return summary


def fill_calculated_cell(dataframe: pd.DataFrame) -> pd.DataFrame:
    """
    Calculates the growth rate for the first row of a DataFrame and adds it as a new cell in the 'Growth Rate' column.

    Args:
        dataframe: A pandas DataFrame with at one row of data and columns 'Thickness' and 'Growth Time' that contain numeric values.

    Returns:
        The input DataFrame with a new cell added to the 'Growth Rate' column.

    Raises:
        TypeError: If the input DataFrame does not meet the above criteria.
    """
    if dataframe.loc[0, 'Thickness'].isnumeric() and dataframe.loc[0, 'Growth Time'].isnumeric():
        dataframe.loc[0, 'Growth Rate'] = pd.to_numeric(dataframe.loc[0, 'Thickness']) / pd.to_numeric(dataframe.loc[0, 'Growth Time'])
    return dataframe


def write_excel_file(file_path: str, dataframe: pd.DataFrame):
    """
    Writes a pandas DataFrame to an Excel file at the specified path.

    Args:
        file_path: The file path where the Excel file should be saved.
        dataframe: The pandas DataFrame to write to the Excel file.

    Returns:
        None
    """
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, sheet_name='Summary', index=False)


def generate_overview(input_file: str):
    """
    Reads data from an experiment file and generates a summary of the data in a new Excel file.

    Args:
        input_file: The name of the experiment file (without extension) to read data from. Must be located in the current working directory.

    Returns:
        None

    Raises:
        FileNotFoundError: If the input file is not found in the current working directory.
    """
    overview = pd.read_excel(input_file, sheet_name="Overview", comment='#', dtype={'Sample': str})
    growth_run = pd.read_excel(input_file, sheet_name="GrowthRun", comment='#')
    afm_reflectance_sem = pd.read_excel(input_file, sheet_name="AFMReflectanceSEM", comment='#')
    hrxrd = pd.read_excel(input_file, sheet_name="HRXRD", comment='#')
    sample_cut = pd.read_excel(input_file, sheet_name="SampleCut", comment='#')

    summary_headers = ['Sample', 'Date', 'Film', 'Substrate', 'Substrate T', 'Carrier Gas', 'Growth Time']

    summary_data: List[Union[str, int, float]] = []
    for overview_header in summary_headers:
        summary_data = fill_column(summary_data, overview, overview_header)

    summary_headers.extend(['Thickness', 'Growth Rate'])
    summary_data = fill_column(summary_data, afm_reflectance_sem, 'Thickness')
    summary_data.append('put calc here')

    for growthrun_header in growth_run:
        for quantity in ['Bubbler Material', 'Gas Cylinder Material', 'Partial Pressure']:
            if quantity in growthrun_header:
                summary_headers.append(growthrun_header)
                if not growth_run[growthrun_header].empty:
                    summary_data.append(growth_run[growthrun_header][0])

    summary_headers.extend(['Phase', 'Collaborator', 'Notes'])
    summary_data = fill_column(summary_data, hrxrd, 'Phase')
    summary_data = fill_column(summary_data, sample_cut, 'Collaborator')
    summary_data.append('put notes here')

    summary_dict = {key: value for key, value in zip(summary_headers, summary_data)}

    filename = os.path.join(os.getcwd(), f"summary_{len(summary_headers)}cols.xlsx")
    new_dataframe = pd.DataFrame(summary_dict, index=[0])
    new_dataframe = fill_calculated_cell(new_dataframe)
    if filename in [os.path.join(os.getcwd(), file_name) for file_name in os.listdir()]:
        existing_dataframe = pd.read_excel(filename, sheet_name='Summary', dtype={'Sample': str})
        num_cols = existing_dataframe.shape[1]
        if num_cols == len(summary_headers):
            updated_dataframe = pd.concat([existing_dataframe, new_dataframe], ignore_index=True)
            write_excel_file(filename, updated_dataframe)
        else:
            write_excel_file(filename, new_dataframe)
    else:
        write_excel_file(filename, new_dataframe)


@click.command()
@click.option(
    '--input-file',
    required=True,
    help='The name of experiment file, without extension. Must be same name of experiment folder'
)
def launch_tool(input_file):
    """
    Main function that reads command line arguments and launches the 'generate_overview' function.

    Args:
        input_file: The name of the experiment file (without extension) to read data from.

    Returns:
        None
    """
    generate_overview(input_file)


if __name__ == '__main__':
    launch_tool().parse()  # pylint: disable=no-value-for-parameter
