## workflow to add a new choice or dropdown field

### Powershell (ps1)
1. 1_deploy_lookup...ps1
    - add content type 
    - add list
    - change default view
2. 2_populate...ps1
    - add options as list items
3. 3_deploy_form...ps1
    - add new lookup field 

### Typescript (ts)
1. SharePointListApi.getListItem()
    - select field/Id, field/Title
    - expand field
2. Utility.convertListItem()
    - let fieldId = this.getIdForDropdown...()
    - return statement, include: field: fieldId
3. Utility.convertFormData()
    - fieldId = formData.field
4. Utility.getOptionsForXFields()
    - const fields = (include field)
5. Utility.getXOptions()
    - case 'field' : list = 'field'; break;

### Components (tsx)
1. Add new component field={field}


## rtp

This is where you include your WebPart documentation.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean 
gulp test 
gulp serve 
gulp bundle -- ship 
gulp package-solution --ship 
