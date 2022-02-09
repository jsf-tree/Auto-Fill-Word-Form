import os
import shutil
import codecs
import math
import warnings
from zipfile import ZipFile
from func_print import final_message, division, section
from import_libs import install_if_nonexistent
install_if_nonexistent('pandas')
install_if_nonexistent('xlrd')
install_if_nonexistent('openpyxl')
import pandas as pd
import numpy as np


os.system("cls")


def read_files(n):
    # Client & Project
    section(n, '1. CLIENT & PROJECT INFORMATION:')
    warnings.simplefilter("ignore")
    cp_data = pd.read_excel(r'input\1_client_project.xlsx').iloc[0:, 1:]
    warnings.simplefilter("default")
    for i in range(len(cp_data)):
        if i not in [0, 6, 7]:
            print(cp_data.iloc[i, 0] + ':', ' ' * (24 - len(cp_data.iloc[i, 0])), cp_data.iloc[i, 1])

    # Sampling data
    section(n, '2. SAMPLING DATA:')
    warnings.simplefilter("ignore")
    sp_data = pd.read_excel(r'input\2_sampling_data.xlsx', index_col=None, skiprows=[0, 1])
    warnings.simplefilter("default")
    print(sp_data)

    # Sampling plan form
    section(n, '3. SAMPLING PLAN:')
    i, path = [0, r'input\\']
    files = os.listdir(path)
    while 'FT-14' not in files[i]:
        i += 1
    ft14 = pd.read_excel(path + files[i], index_col=None, header=None)

    return sp_data, cp_data, ft14


def expected_volume(ft14):
    '''obtains expected volume per sample'''

    def sweeps_df(df, A, stop1, stop2):
        '''Starts at row A col1 goes until B'''
        i_0 = list(df.iloc[:, 0]).index(A)
        while df.iloc[i_0, 0] != stop1:
            i_0 += 1
        i_0 = i_0 + 1
        while str(df.iloc[i_0, 0]) == 'nan':  # Vertically merged cells require search until not NaN
            i_0 += 1
        i_f = i_0
        while stop2 not in df.iloc[i_f, 0]:
            i_f += 1
        return i_0, i_f

    print('Volume/par:\t', end='')
    i0_par, i1_par = sweeps_df(df=ft14, A='INSTRUÇÕES DE COLETA E PRESERVAÇÃO', stop1='Parâmetro', stop2='INSTRUÇÕES')
    par = list(ft14.iloc[i0_par:i1_par, 0])
    vol = list(ft14.iloc[i0_par:i1_par, 4])
    for _ in range(len(par)):
        par[_] = par[_].replace('fingerprint', '').replace(' ', '').replace('dissolvidos', '')
        exec('temporary = ' + vol[_].replace('x', '*').replace('mL', ''), globals())
        vol[_] = temporary

    par_vol = dict(zip(par, vol))
    print(par_vol)

    print('Samples')
    i0_a, i1_a = sweeps_df(df=ft14, A='INSTRUÇÕES DE PREENCHIMENTO DA CADEIA DE CUSTÓDIA GÖRTLER',
                           stop1='ID da Amostra', stop2='INSTRUÇÕES')
    id_sample = list(ft14.iloc[i0_a:i1_a, 0])
    samples_pars = list(ft14.iloc[i0_a:i1_a, 7])

    sample_vol = {'id_sample': [], 'expected_vol': [], 'sampled_vol': []}
    for _ in range(len(id_sample)):
        if 'BC-' not in id_sample[_]:
            e_vol = 0
            for par in list(par_vol.keys()):
                if par in samples_pars[_]:
                    e_vol += int(par_vol[par])
                elif par == 'SVOC/TPH':
                    if 'SVOC' in samples_pars[_] or 'TPH' in samples_pars[_]:
                        e_vol += int(par_vol[par])
            print(id_sample[_] + ': ', e_vol, 'mL')
            sample_vol['id_sample'].append(id_sample[_])
            sample_vol['expected_vol'].append(str(e_vol/1000).replace('.', ','))

    return sample_vol


def process(n, cp_data, sp_data, sample_vol):
    def prepare_dir():
        shutil.rmtree('tmp') if 'tmp' in os.listdir() else os.mkdir('tmp')  # word template dissection here
        shutil.rmtree('output') if 'output' in os.listdir() else os.mkdir('output')  # delete erase files in output

    def unzip():
        with ZipFile('template\FT-23 TEMPLATE.docx', 'r') as zip_ref:
            zip_ref.extractall('tmp')

    def open_word_xmls(cp_data, header_filename, doc_filename):
        header_xml, header_filename, count = [[], header_filename, 0]
        with codecs.open(header_filename, mode='r+', encoding='utf-8') as fileR:
            for row in fileR:
                header_xml.append(row)
        doc_xml, doc_filename, count = [[], doc_filename, 0]
        with codecs.open(doc_filename, mode='r+', encoding='utf-8') as fileR:
            # copies content of document.xml & replace VARX with cp_data
            for row in fileR:
                if count != 0:
                    for _ in range(0, len(cp_data)):
                        row = row.replace('VAR' + str(_), cp_data.iloc[_, 1])
                count += 1
                doc_xml.append(row)
        return header_xml, doc_xml

    def fill_with_sample_data(sp_data, sample_vol, header_xml, doc_xml, header_filename, doc_filename):
        k = 0
        for i in range(sp_data.shape[0]):
            print(''+str(i+1)+':', sp_data.iloc[i, 0])
            # 1. Adjust header var0
            with codecs.open(header_filename, mode='w', encoding='utf-8') as fileW:
                for line in header_xml:
                    line = line.replace('var0', sp_data.iloc[i, 0])
                    fileW.write(line)
            # 2. Adjust doc varx and get vol
            sample_vol['sampled_vol'].append(sp_data.iloc[i, k + 27])
            with codecs.open(doc_filename, mode='w', encoding='utf-8') as fileW:
                for line in doc_xml:
                    for j in range(sp_data.shape[1]):
                        val = sp_data.iloc[i, j]
                        if j == k + 1:
                            line = line.replace('>Sampler1', '>' + str(val))
                        elif j == k + 6:
                            line = line.replace('>Sunny', '>' + str(val))
                        elif j == k + 8:
                            line = line.replace('>No observations', '>' + str(val))
                        elif j == k + 7:
                            line = line.replace('>LNAPL', '>' + str(val))
                        elif j == k + 9:
                            line = line.replace('>SamplingMethod1', '>' + str(val))
                        elif j == k + 10:
                            line = line.replace('>Equipment1', '>' + str(val))
                        elif j == k + 11:
                            line = line.replace('>Equipment2', '>' + str(val))
                        elif j == k + 12:
                            line = line.replace('>Equipment3', '>' + str(val))
                        elif j == k + 13:
                            line = line.replace('>Equipment4', '>' + str(val))
                        else:
                            line = line.replace('>var' + str(j) + '<', '>' + str(val) + '<')
                fileW.write(line)
            # 3. Zip filled sampling form, save to 'output\', rename extension to .docx
            file = r'output\FT-23_' + sp_data.iloc[i, 0]
            file = file.replace(r'/2', '_2')
            file_docx = file + '.docx'
            shutil.make_archive(file, 'zip', 'tmp')
            os.rename(file + '.zip', file_docx)
        return sample_vol

    def report_differences(sp_data, sample_vol):
         df = pd.DataFrame({'Sample ID (FT-14)': sample_vol['id_sample'],
                            'Expected volume (FT-14)': sample_vol['expected_vol'],
                            'FT-25': list(sp_data.iloc[:, 0]),
                            'Sampled volume (FT-25)': sample_vol['sampled_vol']})
         df.to_excel(r'output\differences_report.xlsx', header=True, index=None)

    prepare_dir()
    unzip()
    cp_data = cp_data.drop(index=[0, 6, 7])
    cp_data.index = [*range(len(cp_data))]
    header_filename = 'tmp\word\header1.xml'
    doc_filename = 'tmp\word\document.xml'
    header_xml, doc_xml = open_word_xmls(cp_data, header_filename, doc_filename)
    section(n, '4. FILLING TECHNICAL FORMS nº23...')
    fill_with_sample_data(sp_data, sample_vol, header_xml, doc_xml, header_filename, doc_filename)
    report_differences(sp_data, sample_vol)
    shutil.rmtree('tmp')  # Delete dir: 'tmp\'


def finish(n, sp_data):
    msg = 'Done! ' + str(len(sp_data)) + ' FT-23s were filled and saved to output\!\n' \
                                               'You can double-check now with the other forms: FT-03, 14, 23 and 25!'
    signature = 'Auto-Fill-Word-Form FT-23 v.0\n' \
                'Company 2021\n' \
                '----------------\n' \
                'JSF\n' \
                '12/03/2021\n'
    final_message(msg, signature, n)



def main():
    n = 84
    sp_data, cp_data, ft14 = read_files(n)
    sample_vol = expected_volume(ft14)
    process(n, cp_data, sp_data, sample_vol)
    finish(n, sp_data)


main()
