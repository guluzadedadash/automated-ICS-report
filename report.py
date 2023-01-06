import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import xlsxwriter
from openpyxl import Workbook


@st.cache
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


@st.cache(allow_output_mutation=True)
def load_data(uploaded_file):
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
        except:
            pass

        return df


@st.cache(allow_output_mutation=True)
def vimmm(uploaded_fb_df, uploaded_vim_df):
    via_vim = []
    for ref_num in uploaded_fb_df['Unique Reference Number']:
        if ref_num in list(uploaded_vim_df['Unique Reference Number']):
            via_vim.append(ref_num)
        else:
            via_vim.append('#N/A')
    return via_vim


st.header('App')

uploaded_fb = st.file_uploader('Upload FB data', key='1')
uploaded_pickup = st.file_uploader('Upload pickup data', key='2')
uploaded_user = st.file_uploader('Upload user data', key='3')
uploaded_vim = st.file_uploader('Upload vim data', key='4')

if uploaded_fb and uploaded_vim and uploaded_user and uploaded_pickup:
    uploaded_fb_df = load_data(uploaded_fb)
    uploaded_pickup_df = load_data(uploaded_pickup)
    uploaded_user_df = load_data(uploaded_user)
    uploaded_vim_df = load_data(uploaded_vim)

    final_df = uploaded_fb_df.copy()

    uploaded_fb_df['Document Number'] = uploaded_fb_df['Document Number'].apply(lambda x: str(x).split('.')[0])
    uploaded_fb_df['Unique Reference Number'] = uploaded_fb_df['Document Number'].astype(str) + uploaded_fb_df[
        'Company Code'].astype(str)

    uploaded_vim_df['Document Number'] = uploaded_vim_df['Accounting Document No.'].apply(
        lambda x: str(x).split('.')[0])
    uploaded_vim_df['Unique Reference Number'] = uploaded_vim_df['Accounting Document No.'].astype(str) + \
                                                 uploaded_vim_df['Company Code'].astype(str)
    final_df = uploaded_fb_df.copy()

    uploaded_vim_df['Unique Reference Number'] = uploaded_vim_df['Accounting Document No.'].astype(str) + \
                                                 uploaded_vim_df['Company Code'].astype(str)

    via_vim = vimmm(uploaded_fb_df, uploaded_vim_df)
    final_df['Via VIM'] = via_vim

    # Keeping dataset only NAs in Via VIM
    final_df = final_df[final_df['Via VIM'] == '#N/A']

    #####################

    final_df['User'] = final_df['User Name'].copy()
    final_df['Team'] = final_df['User Name'].copy()

    numbers = uploaded_user_df['nr']
    real_names = uploaded_user_df['name']
    teams = uploaded_user_df['team']

    for number, real_name, team in zip(numbers, real_names, teams):
        final_df['User'] = final_df['User'].replace(str(number), str(real_name))
        final_df['Team'] = final_df['Team'].replace(str(number), str(team))
    ##################
    st.write(final_df)

    value_to_delete_in_user_column = st.multiselect(
        'Select values in `User` column that you want to delete',
        options=final_df['User'].unique(),
        default=['VIM']
    )
    for to_delete in value_to_delete_in_user_column:
        final_df = final_df[final_df['User'] != to_delete]

    value_to_delete_in_team_column = st.multiselect(
        'Select values in `Team` column that you want to delete',
        options=final_df['Team'].unique(),
        default=['OtC', 'Treasury', 'COE DE', 'COE', 'GL', 'Credit and collection']
    )

    for to_delete in value_to_delete_in_team_column:
        final_df = final_df[final_df['Team'] != to_delete]

    to_keep = ['FB60', 'FB65', 'MIRO', 'FB05']
    final_df = final_df[final_df['Transaction Code'].isin(to_keep)]

    final_df['AP Comment'] = [''] * final_df.shape[0]
    final_df['AP Detailed Comment'] = [''] * final_df.shape[0]
    final_df['Justified'] = [''] * final_df.shape[0]

    # for idx, row in enumerate(final_df['User Name']):
    #     if 'polomas' in str(row.lower()):
    #         final_df.loc[idx, 'Justified'] = 'Yes'
    #
    # for idx, row in enumerate(final_df['User Name']):
    #     if 'redwood' in str(row.lower()):
    #         final_df.loc[idx, 'Justified'] = 'Yes'

    we_may_have_these = ['REDWOOD_CET', 'REDWOOD_EST', 'POLOMAS']
    we_have_this = []
    for element in final_df['User Name'].unique():
        if element in we_may_have_these:
            we_have_this.append(element)

    value_to_delete_in_user_name_column = st.multiselect(
        'Select values in `User Name` column that you want to delete',
        options=final_df['User Name'].unique(),
        default=we_have_this
    )

    for to_delete in value_to_delete_in_user_name_column:
        final_df = final_df[final_df['User Name'] != to_delete]

    row_to_delete = []
    for idx, element in enumerate(final_df['Reference']):
        if '_R' in str(element):
            row_to_delete.append(element)

    for to_delete in row_to_delete:
        final_df = final_df[final_df['Reference'] != to_delete]

    final_df['Text'] = final_df['Text'].apply(lambda x: str(x).split('$')[0])
    final_df.loc[final_df['Text'].str.contains('tax|Tax|Correction - incorrect tax|withholding|VAT|WTH|wht|wth|WHT').fillna(False), 'AP Comment'] = 'tax correction'
    final_df.loc[final_df['Text'].str.contains('tax|Tax|Correction - incorrect tax|withholding|VAT|WTH|wht|wth|WHT').fillna(False), 'Justified'] = 'Yes'
    final_df.loc[final_df['Text'].str.contains('Credit|CN|CM|memo|cn|korygująca|kor|do korekty').fillna(False), 'AP Comment'] = 'CN / Credit Note / CM / Credit Memo'
    final_df.loc[final_df['Text'].str.contains('Credit|CN|CM|memo|cn|korygująca|kor|do korekty').fillna(False), 'Justified'] = 'Yes'


    st.write(final_df.shape)
    st.write(final_df.astype(str))
    df_xlsx = to_excel(final_df)
    st.download_button(label='📥 Download in xlsx',
                       data=df_xlsx,
                       file_name='out_of_VIM.xlsx')

