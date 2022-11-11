import * as React from 'react';
import styles from './Crudoperations.module.scss';
import { ICrudoperationsProps } from './ICrudoperationsProps';
import { ICrudoperationsState } from './ICrudoperationsState';
import { escape } from '@microsoft/sp-lodash-subset';
import { ICrudoperations } from './ICrudoperations';  
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  



export default class Crudoperations extends React.Component<ICrudoperationsProps,ICrudoperationsState, {}> {
  constructor(props: ICrudoperationsProps, state: ICrudoperationsState) {  

    super(props);  

  

    this.state = {  

      status: 'Ready',  

      items: []  

    };  

  }  

  public render(): React.ReactElement<ICrudoperationsProps> {
    const items: JSX.Element[] = this.state.items.map((item: ICrudoperations, i: number): JSX.Element => {  
      return (  
        <li>{item.Title} ({item.Id}) </li>  
      );  
    });  
  
    const {
      description,
    
    } = this.props;

    return (
      <div>
      <p >{escape(this.props.listName)}</p>  
                
      <div className={`ms-Grid-row ms-fontColor-white`} >  
        <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>  
          <a href="#"  onClick={() => this.createItem()}>  
            <span >Create an item</span>  
          </a>   
          <a href="#"  onClick={() => this.readItem()}>  
            <span >Read an item</span>  
          </a>  
        </div>  
      </div>  

      <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white `}>  
        <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>  
          <a href="#" onClick={() => this.updateItem()}>  
            <span >Update an item</span>  
          </a>   
          <a href="#"  onClick={() => this.deleteItem()}>  
            <span >Delete an item</span>  
          </a>  
        </div>  
      </div>  

      <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white `}>  
        <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>  
          {this.state.status}  
          <ul>  
            {items}  
          </ul>  
        </div>  
      </div>  

    </div>  
 
);  
}  


private getLatestItemId(): Promise<number> {  
return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {  
this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$orderby=Id desc&$top=1&$select=id`,  
SPHttpClient.configurations.v1,  
{  
  headers: {  
    'Accept': 'application/json;odata=nometadata',  
    'odata-version': ''  
  }  
})  
.then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {  
  return response.json();  
}, (error: any): void => {  
  reject(error);  
})  
.then((response: { value: { Id: number }[] }): void => {  
  if (response.value.length === 0) {  
    resolve(-1);  
  }  
  else {  
    resolve(response.value[0].Id);  
  }  
});  
});  
}   


private createItem(): void {  
this.setState({  
status: 'Creating item...',  
items: []  
});  

const body: string = JSON.stringify({  
'Title': `Test item created by SPFx ReactJS on: ${new Date()}`  
});  

this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`,  
SPHttpClient.configurations.v1,  
{  
headers: {  
'Accept': 'application/json;odata=nometadata',  
'Content-type': 'application/json;odata=nometadata',  
'odata-version': ''  
},  
body: body  
})  
.then((response: SPHttpClientResponse): Promise<ICrudoperations> => {  
return response.json();  
})  
.then((item: ICrudoperations): void => {  
this.setState({  
status: `Item '${item.Title}' (ID: ${item.Id}) successfully created`,  
items: []  
});  
}, (error: any): void => {  
this.setState({  
status: 'Error while creating the item: ' + error,  
items: []  
});  
});  
}  

private readItem(): void {  
this.setState({  
status: 'Loading latest items...',  
items: []  
});  

this.getLatestItemId()  
.then((itemId: number): Promise<SPHttpClientResponse> => {  
if (itemId === -1) {  
  throw new Error('No items found in the list');  
}  

this.setState({  
  status: `Loading information about item ID: ${itemId}...`,  
  items: []  
});  
return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemId})?$select=Title,Id`,  
  SPHttpClient.configurations.v1,  
  {  
    headers: {  
      'Accept': 'application/json;odata=nometadata',  
      'odata-version': ''  
    }  
  });  
})  
.then((response: SPHttpClientResponse): Promise<ICrudoperations> => {  
return response.json();  
})  
.then((item: ICrudoperations): void => {  
this.setState({  
  status: `Item ID: ${item.Id}, Title: ${item.Title}`,  
  items: []  
});  
}, (error: any): void => {  
this.setState({  
  status: 'Loading latest item failed with error: ' + error,  
  items: []  
});  
});  
}  

private updateItem(): void {  
this.setState({  
status: 'Loading latest items...',  
items: []  
});  

let latestItemId: number = undefined;  

this.getLatestItemId()  
.then((itemId: number): Promise<SPHttpClientResponse> => {  
if (itemId === -1) {  
  throw new Error('No items found in the list');  
}  

latestItemId = itemId;  
this.setState({  
  status: `Loading information about item ID: ${latestItemId}...`,  
  items: []  
});  
  
return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${latestItemId})?$select=Title,Id`,  
  SPHttpClient.configurations.v1,  
  {  
    headers: {  
      'Accept': 'application/json;odata=nometadata',  
      'odata-version': ''  
    }  
  });  
})  
.then((response: SPHttpClientResponse): Promise<ICrudoperations> => {  
return response.json();  
})  
.then((item: ICrudoperations): void => {  
this.setState({  
  status: 'Loading latest items...',  
  items: []  
});  

const body: string = JSON.stringify({  
  'Title': `Updated Item ${new Date()}`  
});  

this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,  
  SPHttpClient.configurations.v1,  
  {  
    headers: {  
      'Accept': 'application/json;odata=nometadata',  
      'Content-type': 'application/json;odata=nometadata',  
      'odata-version': '',  
      'IF-MATCH': '*',  
      'X-HTTP-Method': 'MERGE'  
    },  
    body: body  
  })  
  .then((response: SPHttpClientResponse): void => {  
    this.setState({  
      status: `Item with ID: ${latestItemId} successfully updated`,  
      items: []  
    });  
  }, (error: any): void => {  
    this.setState({  
      status: `Error updating item: ${error}`,  
      items: []  
    });  
  });  
});  
}   



private deleteItem(): void
{  
if (!window.confirm('Are you sure you want to delete the latest item?'))
{  
return;  
}  

this.setState({  
status: 'Loading latest items...',  
items: []  
});  

let latestItemId: number = undefined;  
let etag: string = undefined;  
this.getLatestItemId()  
.then((itemId: number): Promise<SPHttpClientResponse> => {  
if (itemId === -1) {  
  throw new Error('No items found in the list');  
}  

latestItemId = itemId;  
this.setState({  
  status: `Loading information about item ID: ${latestItemId}...`,  
  items: []  
});  

return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${latestItemId})?$select=Id`,  
  SPHttpClient.configurations.v1,  
  {  
    headers: {  
      'Accept': 'application/json;odata=nometadata',  
      'odata-version': ''  
    }  
  });  
})  
.then((response: SPHttpClientResponse): Promise<ICrudoperations> => {  
etag = response.headers.get('ETag');  
return response.json();  
})  
.then((item: ICrudoperations): Promise<SPHttpClientResponse> => {  
this.setState({  
  status: `Deleting item with ID: ${latestItemId}...`,  
  items: []  
});  

return this.props.spHttpClient.post(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${item.Id})`,  
  SPHttpClient.configurations.v1,  
  {  
    headers: {  
      'Accept': 'application/json;odata=nometadata',  
      'Content-type': 'application/json;odata=verbose',  
      'odata-version': '',  
      'IF-MATCH': etag,  
      'X-HTTP-Method': 'DELETE'  
    }  
  });  
})  
.then((response: SPHttpClientResponse): void => {  
this.setState({  
  status: `Item with ID: ${latestItemId} successfully deleted`,  
  items: []  
});  
}, (error: any): void => {  
this.setState({  
  status: `Error deleting item: ${error}`,  
  items: []  
});  
});  
}  

 

}
