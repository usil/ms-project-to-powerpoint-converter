# Power point with Python

  :warning: Avoid leaving open the files that will be used

## Requirements

- Python >= 3.7 [installation](https://www.digitalocean.com/community/tutorials/install-python-windows-10)
- Install PIP on Windows [installation](https://www.geeksforgeeks.org/how-to-install-pip-on-windows/) 
[others](https://stackoverflow.com/questions/23708898/pip-is-not-recognized-as-an-internal-or-external-command)

### Libraries

- pywin32 [documentation](https://pypi.org/project/pywin32/)


## Parameters
Shipping parameters in the console.

| Parameter| Description| Example|
| ---| ---| ---|
| --source_folder or -sf| Source folder origin| "C:/docs"|
| --index_slides or -is| Index slides to update| "1,2"|
| --base_node or -bn| Base node| "Project > System 1"

Command to list required arguments

```python
py main.py -h
```

## Manual start
```python
pip install -r requirements.txt

py main.py -sf "source folder" -is "1,2" -bn "Project > System 1"
```

## Steps
1. Create a folder containing the Ms project files, with the following parameters
  - MS project file:
  - Power point file: 

    Project example: [docs.zip](https://github.com/usil/ms-project-to-powerpoint-converter/files/12501687/docs.zip)

  :warning: File names must not contain commas or special characters.

  :warning: There must be a single base file.

  <p align="center">
    <img src="https://github.com/usil/ms-project-to-powerpoint-converter/assets/77288944/c3db736b-8bb3-46a9-9ae4-165556e4c9f8" width="100%">
  </p>
  <p align="center">
    <img src="https://github.com/usil/ms-project-to-powerpoint-converter/assets/77288944/88e62629-7cbd-44ba-895e-0e637839fe75" width="100%">
  </p>
  <p align="center">
    <img src="https://github.com/usil/ms-project-to-powerpoint-converter/assets/77288944/d366a8b2-45a8-408a-8525-80f66bbe92d1" width="100%">
  </p>

2. Run command

:warning: Remember to close all files .mpp before execute this script

```python
pip install -r requirements.txt

py main.py -sf "source folder" -is "1,2" -bn "Project > System 1"
```


3. Result 


## Contributors

<table>
  <tbody>
    <td>
      <img src="https://avatars.githubusercontent.com/u/77288944?v=4" width="100px;"/>
      <br />
      <label><a href="https://github.com/madeliyricra">Madeliy Ricra</a></label>
      <br />
    </td>  
  </tbody>
</table>

## Documentation
- https://stackoverflow.com/questions/71430344/update-links-of-powerpoint-using-win32com-python
- https://mhammond.github.io/pywin32/
- https://stackoverflow.com/questions/71430344/update-links-of-powerpoint-using-win32com-python
- https://stackoverflow.com/questions/55227428/opening-a-powerpoint-presentation-saving-as-pdf-and-closing-the-application-usi
- https://stackoverflow.com/questions/55942773/how-to-set-title-to-a-powerpoint-slide-using-win32com-client
- https://stackoverflow.com/questions/73233231/how-to-add-buttons-for-macro-by-using-xlwings-or-pywin32