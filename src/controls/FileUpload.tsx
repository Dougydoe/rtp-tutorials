import * as React from 'react';
import { IFileUploadProps, IFileUploadState } from '../interfaces';
import { DropzoneComponent } from 'react-dropzone-component';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import file_styles from './FileUpload.module.scss';
import { Label, DetailsList,  IColumn, SelectionMode } from 'office-ui-fabric-react';
import { FileApi } from '../utils/';
import styles from './FieldStyles.module.scss';


export class FileUpload extends React.Component<IFileUploadProps, IFileUploadState> {
  
  constructor(props: IFileUploadProps) {
    super(props);

    this.state = {      
      digest: null,
      attachments: [],      
    };
  }
  
  /**
   * @description updates state with digest and if there are existing attachments updates state with those too
   */
  public async componentDidMount():Promise<void> {

    this.props.context['serviceScope'].whenFinished(() => {
      const digestCache: IDigestCache = this.props.context["serviceScope"].consume(DigestCache.serviceKey);      
      digestCache.fetchDigest(this.props.context.pageContext.web.serverRelativeUrl).then((digest: string): void => {            
        this.setState({digest: digest});
      });  
    });
    
    if(this.props.formContext.mode == "edit") {
      let attachments = await FileApi.getAttachments(this.props.context.pageContext.web.serverRelativeUrl, this.props.formContext.formGuid, this.props.formData);
      this.setState({attachments: attachments});
    }

    let error:string = this.getFooterText();
    if (error) {
      this.props.onError(this.props.field);
    } else if (!error) {
      this.props.onError(this.props.field, true);
    }

  }

  public shouldComponentUpdate(nextProps):boolean {
    // * I need to update when formData changes else I do not know if an attachment has been added
    /* if (this.props.formData != nextProps.formData) {
        return false;
    } */ 
    return true;
  }

  public componentDidUpdate() {
    // console.log('componentDidUpdate field: ' + this.props.field);        
    // do validation 
    let error:string = this.getFooterText();
    if (error) {
        this.props.onError(this.props.field);
    } else if (!error) {
        this.props.onError(this.props.field, true);
    }
  }

  private fieldRequired = ():boolean => {
    if (this.props.validation && this.props.validation[this.props.field]) {
        const val:any = this.props.validation[this.props.field];
        if (val.validateWhen == null || val.validateWhen(this.props.formData)) {                                             
            return val.required;
        }
    }
    return false;
  }

  private getFooterText = ():string => {
    let value = null;
    let footer_text:string = "";
    let fieldRequired:boolean = this.fieldRequired();

    //get field value
    if (this.props.formData[this.props.field]) {
        value = this.props.formData[this.props.field];
    }
    // * check if attachments already exist if its an existing form 
    if (this.state.attachments.length > 0) {
      return footer_text;
    } else {
      // * there are no existing attachments saved in SharePoint
      // * check if field is required and has no attachments
      if (fieldRequired) {
        if (!this.props.formData[this.props.field] || (this.props.formData[this.props.field].length == 0)) {
          footer_text = "Please add an attachment";                        
        }
      }
      return footer_text;
    }

    
  }
  
  public render(): React.ReactElement<IFileUploadProps> {
    const footer_text:string = this.getFooterText();
    let context = this.props.context;
    let formGuid = this.props.formContext.formGuid;    
    let myDropzone;
    const fieldRequired:boolean = this.fieldRequired();


    let componentConfig = {
      postUrl: context.pageContext.web.absoluteUrl
    };

    var djsConfig = {
      headers: {
        "X-RequestDigest": this.state.digest
      },
      addRemoveLinks:true
    };

    let eventHandlers = {      
      init: (dz) => {       
       myDropzone=dz;
      },
      /**
       * @description update state.formData.Attachments to show a null value for the removed file 
      */
      removedfile: (file) => {
        this.props.onUpdate(file.name, null, true);        
      },
      /**
       * @description update state.formData.Attachments with the newly added file 
      */
      processing: (file) => {                        
        this.props.onUpdate(file.name, file, true);        
      },        
      error: (file) => {    
        console.log(`Error, file '${file.name}' was not added.`);
      }
     };

     const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'Name',
        fieldName: 'Name',
        minWidth: 210,
        maxWidth: 350,
        isResizable: true,
        isPadded: true,
        onRender: (item: any) => {
          return <a target="_blank" href={item.LinkingUri? item.LinkingUri : item.ServerRelativeUrl}>{item.Name}</a>;
        }
      },
      {
        key: 'column2',
        name: 'Date Modified',
        fieldName: 'TimeLastModified',
        minWidth: 70,
        maxWidth: 150,
        isResizable: true,
        data: 'number',
        isPadded: true,
        onRender: (item: any) => {
          let date:Date = new Date(item.TimeLastModified);
          return <span>{date.toLocaleString("en-gb")}</span>;
        },
      },
      {
        key: 'column3',
        name: 'Date Created',
        fieldName: 'TimeCreated',
        minWidth: 70,
        maxWidth: 150,
        isResizable: true,
        data: 'number',
        isPadded: true,
        onRender: (item: any) => {
          let date:Date = new Date(item.TimeCreated);
          return <span>{date.toLocaleString("en-gb")}</span>;
        },
      }
    ];


    return (      
      <div className={ file_styles.row }>
        <Label required={fieldRequired}>{this.props.label}</Label>
        <DropzoneComponent eventHandlers={eventHandlers} djsConfig={djsConfig} config={componentConfig}>
          <div className="dz-message">Drop files here or click to upload.</div>
        </DropzoneComponent>
        <span className={styles.dsErrorLabel}>{footer_text}</span>
        <br/>
        {(this.state.attachments.length > 0) ? 
          <DetailsList 
            items={this.state.attachments}          
            columns={columns}   
            selectionMode={SelectionMode.none}       
          /> : null }
      </div>      
    );
  }
}
