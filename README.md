## Auto-Soil-Profile-Drawer
#### Wondering how is it at work?
> Check on my YouTube -> https://youtu.be/pXsSaKJ_2TY
#### Why?
> Filling word forms was too time-consuming in the company. 
#### How was the problem solved?
> By setting a word doc template with tag-words, and developing a script to replace them from the data of two .xlsx, which are faster to fill.
#### What did I learn?
> I learned how to optimize filling word docs.
---
#### FIRST TIME USING IT?
Make sure you have a python.exe set as an environmental variable in PATH
```
Control panel > System > Advanced system settings > Environment Variables
Under "User variables", select "Path" > Edit > New > 
Type the path to your python.exe
Move this path to the top of the list
```
---
#### INSTRUCTIONS:
1. Fill the .xlsx in "input\" with data, add a sampling plan form
2. Run "run.bat"
3. Volumes will be checked and a words doc will be filled by sample.
4. Files will appear in "output\"
- When filling both xlsx, mind the formatting: point as decimal sepator


---
#### EXPLANATION TO FILES AND FOLDERS:
##### Files:
- **input\1_client_project.xlsx**\
_Generic information (client and project)_
- **input\2_sampling_data.xlsx**\
_Measurements taken in the field during sampling_
- **input\FT-14.xls**\
_Measurements taken in the field during sampling_
- **template\FT-23 TEMPLATE.docx**\
_The basic template for the forms to be filled_
- **1_run.bat**\
_Used to run python_
- **LICENSE**\
_License in GitHub_
- **README.md**\
_This readme_
##### Folders
- **output\\**\
_directory where filled forms and reports are saved to_
- **template\\**\
_directory where .py and teamplate are stored_
