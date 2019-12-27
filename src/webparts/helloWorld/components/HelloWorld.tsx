import * as React from 'react';
import axios from 'axios';
import pnp from "sp-pnp-js"; 
import { PageContext } from '@microsoft/sp-page-context'
import ReactFileReader from 'react-file-reader';    
import {Icon} from 'office-ui-fabric-react/lib/Icon'; 
import {DefaultButton } from 'office-ui-fabric-react/lib/Button'; 
import {Fabric} from 'office-ui-fabric-react/lib/Fabric';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps } from 'office-ui-fabric-react/lib/Dropdown';
import { TextField, MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import {Stack,IStackProps} from 'office-ui-fabric-react/lib/Stack';
import { IPersonaProps, Persona } from 'office-ui-fabric-react/lib/Persona';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import {
  CompactPeoplePicker,
  IBasePickerSuggestionsProps,
  IBasePicker,
  ListPeoplePicker,
  NormalPeoplePicker,
  ValidationState
} from 'office-ui-fabric-react/lib/Pickers';

export default class HelloWorld extends React.Component<IHelloWorldProps, any> {
  
  
  constructor(props:any){
      super(props);
     
      this.state={listItems:[],user:[],formItem:{}};
      
  }
  
//  private currentPageUrl = this.props.site ;
//  private spWeb = new Web(this.context.pageContext.web.absoluteUrl); 
  private columns = [
    { key: 'column1', name: 'title', fieldName: 'title',type:'string', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column2', name: 'assigned to', fieldName: 'assignedTo',type:'string', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column3', name: 'desc', fieldName: 'desc',type:'string', minWidth: 100, maxWidth: 200, isResizable: true },
     { key: 'column4', name: 'files', fieldName: 'files',type:'file', minWidth: 100, maxWidth: 200, isResizable: true },
  ];

  public componentDidMount(){
   // console.log('currentPageUrl', this.spWeb);
    let items=[];
    let users=[];
    
    let attachmentfiles: string = "";
 
    pnp.sp.web.lists.getByTitle("taskList").items
    .select("Id,Title,Note,AssignedTo/Title,Attachments,AttachmentFiles")
    .top(3).orderBy("Created",false)
    .expand("AttachmentFiles","AssignedTo/Id")
    .filter('Attachments eq 1')
    .get().then((response) => {
     console.log('response',response);
      response.forEach((listItem: any) => {
        var assignedStr=listItem.AssignedTo?listItem.AssignedTo.reduce((accVal,curVal,idx)=> idx==0? curVal.Title:accVal +',' +curVal.Title,''):'';
        items.push({
          title: listItem.Title,
          assignedTo:assignedStr,
          desc:listItem.Note,
          files: listItem.AttachmentFiles
        });
      });    
      this.setState({listItems:items});
     // this.renderData(attachmentfiles);
    })
    pnp.sp.web.siteUsers.get().then((rsp)=> {
     console.log('userssss',rsp);
      rsp.forEach(itm=>{
        users.push({key:itm.Id,text:itm.Title});
      });
   
      console.log(users);
      this.setState({users:users});
  });
   
  }
  private fileChange=(e)=>{

    console.log('fileeeee',e.target.value +'..'+ e.files.fileList);
  }
 private renderItemColumns(){

  return  this.columns.map((column)=>{
    return {
    ...column,

    type:column.type,
    onRender: (item?: any, index?: number) =>( <span>
   
      {(() => { 
               switch (column.type) {
                  
                    case 'file':
                    return (
                      <span>{item.files.map(file=>{
                         
                         return <a href={file.ServerRelativeUrl}>{file.FileName}</a>
                         })}</span>
                    )
                    default:
                    return (
                      item[column.fieldName]
                    )
                  }
              })()}                                                         
      </span>)
  }
    })
    
 }
 private HandleUploadedFiles=(files)=>{  
  let fileSave=[];
  let value='';
  

    // for(var i=0;i<files.length;i++){
    //   let Title = files[i].name.substring(0,files[i].name.lastIndexOf('.'));
    //   let fileExtention = files[i].name.substring(files[i].name.lastIndexOf('.')+1, files[i].name.length) || Title;
    //   let FileName =Title+ '.' + fileExtention
    //   let Content=e.base64[i].split('base64,')[1];
    //     fileSave.push({name:FileName,content:Content});
        
    //     }

    var fileExtension = this.GetFileExtension(files.fileList[0].name);    
            
            var fileInternalName = files.fileList[0].name.substr(0, files.fileList[0].name.lastIndexOf('.'));    
            var fileName = fileInternalName + "." + fileExtension;    
            var myBlob = this.Base64ToArrayBuffer(files.base64);  
       this.setState({formItem:{...this.state.formItem,attachments:{name:fileName,content:myBlob}}})
}

public GetFileExtension(filename) {    
  try{    
      return (/[.]/.exec(filename)) ? /[^.]+$/.exec(filename)[0] : undefined;    
  }catch(error){    
      console.log("Error in Get File Extension  " + error);    
  }            
}   
  public Base64ToArrayBuffer(base64) {    
  try {    
      var binary_string = window.atob(base64.split(',')[1]);    
      var len = binary_string.length;    
      var bytes = new Uint8Array(len);    
      for (var i = 0; i < len; i++) {    
          bytes[i] = binary_string.charCodeAt(i);    
      }    
      return bytes.buffer;    
  } catch (error) {    
      console.log("Error in Base64ToArrayBuffer " + error);    
  }    
}  
private changeSelectedUsers=(option)=>{
  console.log('optionsssss',option)
  let newselectedkeys=this.state.formItem.users?[...this.state.formItem.users]:[];
  if (option.selected) {
    newselectedkeys.push(option.key);
  } else {
    // Remove the item from the selected keys list
    const itemIdx = newselectedkeys.indexOf(option.key as string);
    if (itemIdx != -1) {
      newselectedkeys.splice(itemIdx, 1);
    }
  }
 this.setState({formItem:{...this.state.formItem,users:newselectedkeys}})

}  
//
private handleOnChange(event): void{
    this.setState({formItem: { ...this.state.formItem, [event.target.name]: event.target.value}} )
}
private saveItem():void{


   let item=this.state.formItem;
   let assignedTo=this.state.formItem.users.map((a) => (a.Id));
    pnp.sp.web.lists.getByTitle("taskList").items.add({
    Title: this.state.formItem.title,
    Note:this.state.formItem.note,
    AssignedToId:{results:this.state.formItem.users}
       // allows a single lookup value
    // MultiLookupFieldId: { 
    //     results: [ 1, 56 ]  // allows multiple lookup value
    // }
}).then((r) => {
  console.log("item id is:", r.data.Id);
  console.log('attachments',this.state.formItem.attachments);
  if(this.state.formItem.attachments)
      r.item.attachmentFiles.add(this.state.formItem.attachments.name,this.state.formItem.attachments.content).then(v => {
        console.log(v);
    }).catch(e=>alert('file failed to save'))
    alert('Item saved successfully');

   }).catch(error=>{
    alert('Item failed to save');
    console.log('error',error);
   }) 
}

  public render(): React.ReactElement<IHelloWorldProps> {
    const columnProps: Partial<IStackProps> = {
      tokens: { childrenGap: 15 },
      styles: { root: { width: '80%', padding: 10} }
    };
    
    let items=this.state.listItems;
    let users=this.state.users;
    return (
      <Fabric>
         <DetailsList 
         items={items}
         columns={this.renderItemColumns()}
        
         setKey="set"
         layoutMode={DetailsListLayoutMode.justified}
        isHeaderVisible={false}
      
         selectionPreservedOnEmptyClick={true}
         ariaLabelForSelectionColumn="Toggle selection"
         ariaLabelForSelectAllCheckbox="Toggle selection for all items"
         checkButtonAriaLabel="Row checkbox"
         
       />
          <Stack {...columnProps} >
          <div style={{marginTop:'5%',fontSize: 'x-large'}}>Enter new Item</div>
          <Stack.Item >Title:  <TextField name='title' value={this.state.formItem.title} onChange={e=>this.handleOnChange(e)}  /> </Stack.Item>
          <Stack.Item >Assign To:<Dropdown placeholder="Select option" multiSelect={true}
          options={users} 
          //  onChanged={ (option) => this.setState({formItem: { ...this.state.formItem, users:option}}) } 
          onChanged={option=>this.changeSelectedUsers(option)}
           ></Dropdown></Stack.Item >
          <Stack.Item>Note: <TextField name='note' multiline  value={this.state.formItem.note} onChange={e=>this.handleOnChange(e)}/></Stack.Item>
          <Stack.Item>Attachment:   
          {/* <input type='file'  onChange={(e)=>this.HandleUploadedFiles(e)}/> */}
          <div id="fileUpload1">    
          <ReactFileReader   handleFiles={(e)=>this.HandleUploadedFiles(e)} base64={true} >
          <input type='button' value='Upload'/>
          <span  style={{marginLeft:'3%'}}>{this.state.formItem.attachments?this.state.formItem.attachments.name:'' }</span>    
          </ReactFileReader> 
          </div> 
          
          </Stack.Item >  
          
        <div><DefaultButton text="Save"  onClick={()=>this.saveItem()} /></div>
          </Stack>
      </Fabric>
      
      )
  }
}
           
   
      
    
