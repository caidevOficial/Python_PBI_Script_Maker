<table>
    <tr>
        <td align='center'>
            <img alt="Logo SEA Team" src="./assets/img/banner.png?raw=true" height="145px" />
        </td>
    </tr>
    <tr>
        <td align="center">
            <img alt="Logo Python" src="https://github.com/devicons/devicon/raw/master/icons/python/python-original.svg?raw=true" height="155px" />
        </td>
    </tr>
</table></br>

---

# SEA Power BI Script Maker
#### Developed by: _Facundo Falcone_ <amilcar.f.falcone@accenture.com>
<br>

### Important: You`ll need Python 3.10.X

### How to use:

As a first step, install the necessary libraries to run the program, for this you must execute the command in the console:

If you use windows:
```bash
py -m pip install -r requirements.txt
```

If you use UNIX:
```bash
python3 -m pip install -r requirements.txt
```

This command will install all the necessary libraries to run the program without problems.

As a second step, you should edit the file [source_pbi_fields.xlsx](source_pbi_fields.xlsx) **Without erasing the red headers**, from line 2 onwards, putting the technical names of the fields in the first column and their functional names in the second column.

<table align='center'>
    <thead>
        <theader>
            <center>xlsx file</center>
        </theader>
    </thead>
    <tbody>
        <tr>
            <td>
                <img src='./assets/img/config_doc.png'>
            </td>
        </tr>
    </tbody>
</table>

As a third step, in the GUI you will have a textbox where you will put the name of the table for which you want to write the script and select from the combobox the dataset to which that table belongs. Once done, click on the blue button.
This action will create a file with extension _.vba_ which will contain the script to copy and paste it into **Power BI**.

<table align='center'>
    <thead>
        <theader>
            <center>Main GUI</center>
        </theader>
    </thead>
    <tbody>
        <tr>
            <td>
                <img src='./assets/img/app_gui.png'>
            </td>
        </tr>
    </tbody>
</table>

The program will create a file like: **_dataset_name.table_name.vba_** which will contain the script to copy in the advanced settings of **Power BI**

```js
    let
        Source = GoogleBigQuery.Database([BillingProject = ProjectID, UseStorageApi = false]),
        Navigation = Source{[Name = DatalakeID]}[Data],
        #"Navigation 1" = Navigation{[Name = "sea_procurement_196220_in", Kind = "Schema"]}[Data],
        #"Navigation 2" = #"Navigation 1"{[Name = "sea_sap_fieldglass_sow_in", Kind = "Table"]}[Data],
        #"Renamed columns" = Table.RenameColumns(#"Navigation 2", {{"business_unit", "Business Unit"}, {"business_unit_code", "Business Unit Code"}, {"country_region", "Country Region"}, {"vendor_id", "Vendor ID"}, {"supplier_code", "Supplier Code"}, {"supplier", "Supplier"}, {"supplier_invitation_create_time", "Supplier Invitation Create Time"}, {"supplier_invitation_accept_time", "Supplier Invitation Accept Time"}, {"supplier_invitation_status", "Supplier Invitation Status"}, {"supplier_status", "Supplier Status"}, {"count_of_statement_of_work", "Count of SOWs"}, {"supplier_onboarding_time", "Supplier Onboarding Time"}, {"submit_for_supplier_review_date", "Submit for Supplier Review Date"}, {"approval_sequence", "Approval Sequence"}, {"activation_status", "Activation Status"}, {"supplier_response_time", "Supplier Response Time (hours)"}, {"number_of_sites", "Number of Sites"}})
    in
        #"Renamed columns"
    
```

<table align='center'>
    <tr align='center'>
        <h2 align='center'>Technologies used. 📌</h2>
        <td>
            <a href="https://www.python.org/downloads/"><img alt="Pyhton Logo" src="https://github.com/caidevOficial/Logos/blob/master/Lenguajes/py_logo1_1.png?raw=true" width="50px" height="50px" /></a>
        </td>
        <td><center>Python</center></td>
    </tr>
    <tr align='center'>
        <td>
            <a href="https://pandas.pydata.org/"><img alt="Pandas Logo" src="https://upload.wikimedia.org/wikipedia/commons/thumb/e/ed/Pandas_logo.svg/1200px-Pandas_logo.svg.png?raw=true" height="50px" /></a>
        </td>
        <td><center>Pandas</center></td>
    </tr>
    <tr align='center'>
        <td>
            <a href="https://www.pygame.org/news"><img alt="Pygame Logo" src="https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Pygame_logo.gif/640px-Pygame_logo.gif?raw=true" height="50px" /></a>
        </td>
        <td><center>Pygame</center></td>
    </tr>
    <tr align='center'>
        <td>
            <a href="https://code.visualstudio.com/"><img alt="VSCode Logo" src="https://github.com/caidevOficial/Logos/blob/master/Lenguajes/visual-studio-code.svg?raw=true" height="50px" /></a>
        </td>
        <td><center>VSCode</center></td>
    </tr>
</table>
<br><br><br>