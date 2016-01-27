# CopyFieldRef

- Will take care about adjusting field definitions in sharepoint to the right format for references.
- Just select fields for which you need to create reference, push keybind for macro and reference will be in your clipboard.
- You can select as many field definitions as you like and it's enough if you'll select just *ID* and *Name*.
- If number of selected *IDs* and *Names* wont't be even you'll endup with ERROR 9.
- If you prefer your references to be on one line change *OneLiner* to True

### Example
* **Field definition:**
    ```sh
    <Field ID="{A42B9CD0-9517-4BAA-9417-3CFF52E5BD21}"
           Name="SelectPeasant"
           StaticName="SelectPeasant"
           DisplayName="Select peasant"
           Group="Custom Fields"
           Type="User"
           Required="FALSE" />
    ```
* **Adjusted reference:**
    * oneLiner = False
        ```sh
        <FieldRef ID="{A42B9CD0-9517-4BAA-9417-3CFF52E5BD21}"
    			  Name="SelectPeasant" />
    	```
	* oneLiner = True
    	```sh
    	<FieldRef ID="{A42B9CD0-9517-4BAA-9417-3CFF52E5BD21}" Name="SelectPeasant" />
    	```
