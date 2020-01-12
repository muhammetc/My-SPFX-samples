import * as React from 'react';
import styles from './SpCrud.module.scss';
import { ISpCrudProps } from './ISpCrudProps';
import { ISpCrudState } from './ISpCrudState';
import { IListItem } from './IListItem';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { confirmAlert } from 'react-confirm-alert';
import Modal from 'react-awesome-modal';

import 'react-confirm-alert/src/react-confirm-alert.css';



export default class SpCrud extends React.Component<ISpCrudProps, ISpCrudState> {

  constructor(props: ISpCrudProps) {
    super(props);

    this.state = {
      isVisible: false,
      status: "",
      Items: [],
      itemTitle: "",
      itemId: "",
      isAscedingSort:false,
    };
  }
  public componentDidMount(): void {
    if (this.props.listName) {
      this._getListItems();
    }

  }
  changeInput = (e) => {

    console.log(e);
    this.setState({
      itemTitle: e.target.value
    })
  }

  onSortListItem=()=>{
    debugger;
    console.log(this.state.isAscedingSort);
    this.setState({
      isAscedingSort:!this.state.isAscedingSort
    })
    this._getListItems();
/// Veri ekleme ve güncelleme sırasında verilerin sıralamsı değiştiği için  bu yöntem kullanılmadı
   /*  console.log(this.state.isAscedingSort);
    if(this.state.isAscedingSort){
      this.setState(prevState => {
        this.state.Items.sort((a, b) => (parseInt(a.Id) - parseInt(b.Id)))
    });
    }
    else{
      this.setState(prevState => {
        this.state.Items.sort((a, b) => (parseInt(b.Id) - parseInt(a.Id)))
    });
    } */
  }
  onShowEditPanel = (itemId, itemTitle) => {
    let title = "";
    if (itemTitle) {
      title = itemTitle;
    }
    this.setState({
      itemTitle: title,
      isVisible: true,
      itemId: itemId
    })
  }

  onShowAddPanel = () => {
    this.setState({
      itemTitle: "",
      isVisible: true,
      itemId: "",
      status: ""
    })
  }

  onSaveOrUpdateListItem = () => {

    if (this.state.itemTitle == "") {
      this.setState({
        status: "Title is not blank"
      })
    }
    else {
      if (this.state.itemId !== "") {

        this._updateListItem();
      }
      else {
        this._addListItem();
      }
    }
  }


  onDeleteItem = (e) => {
    this.setState({

      itemId: e
    })
    confirmAlert({
      customUI: ({ onClose }) => {
        return (
          <div className='custom-ui'>
            <h1>Are you sure?</h1>
            <p>You want to delete this file?</p>
            <button onClick={onClose}>No</button>
            <button
              onClick={() => {
                this._deleteListItem();
                onClose();
              }}
            >
              Yes, Delete it!
            </button>
          </div>
        );
      }
    });
  }
  public render(): React.ReactElement<ISpCrudProps> {

    return (
      <div className={styles.spCrud}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <p className={styles.title}>SharePoint Content!</p>

              <a href="#" className={styles.button} onClick={() => this.onShowAddPanel()}>
                <span className={styles.label}>Add List Item</span>
              </a>

            </div>
          </div>

          <div className={styles.row}>
            <i className={this.state.isAscedingSort? 'ms-Icon ms-Icon--Descending' : 'ms-Icon ms-Icon--Ascending'}
             onClick={this.onSortListItem}
             style={{fontSize:"24px", marginBottom:"5px"}}
             ></i>
            <ul className={styles.list}>
              {this.state.Items &&
                this.state.Items.map((list) =>
                  <li key={list.Id} className={styles.item}>
                    <strong>Id:</strong> {list.Id}, <strong>Title:</strong> {list.Title}
                    <i className="ms-Icon ms-Icon--Edit" style={{ marginLeft: "20px" }} onClick={() => this.onShowEditPanel(list.Id, list.Title)}></i>
                    <i className="ms-Icon ms-Icon--Delete" style={{ marginLeft: "10px" }} onClick={() => this.onDeleteItem(list.Id)}></i>
                  </li>
                )}

            </ul>
          </div>


          <Modal
            visible={this.state.isVisible}
            width="500"
            height="130"
            effect="fadeInUp"

          >

            <div className="row" id='dvUpdatePane' style={{ backgroundColor: "#005A9E", borderStyle: "solid", borderColor: "darkgray", borderWidth: "2px" }}>
              <div className="col-md-12 mb-3" style={{ paddingTop: "20px" }}>
                <div>

                  <div>
                    <input type="text" id="input" className="form-control"
                      placeholder="Enter List Item Title"
                      value={this.state.itemTitle} onChange={this.changeInput}
                    ></input>
                    <br />
                    <div className="row" >
                      <div className="col-md-6">

                      </div>
                      <div className="col-md-6">
                      <a href="#" className="btn btn-primary" onClick={this.onSaveOrUpdateListItem}>
                          <span className={styles.label} style={{padding:"0px 15px 0px 15px"}}>Save</span>
                        </a>
                        <a href="#" className="btn btn-danger" style={{marginLeft:"15px"}} onClick={() => this.setState({isVisible:false})}>
                          <span className={styles.label} style={{padding:"0px 15px 0px 15px"}}>Close</span>
                        </a>
                      </div>
                    </div>
                  </div>
                  <div>
                    <span className="text-white">{this.state.status}</span>
                  </div>
                </div>
              </div>
            </div>
          </Modal>


        </div>

      </div>
    );
  }


  private _getListItems(): void {
    this.setState({


      itemTitle:""
    });
    debugger;
    let sort="";
    if(this.state.isAscedingSort){
      sort="&$orderby=Id asc"
    }
    else{
      sort="&$orderby=Id desc"
    }
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getById('${this.props.listName}')/items?$select=Title,Id${sort}`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse): Promise<{ value: IListItem[] }> => {
        return response.json();
      })
      .then((response: { value: IListItem[] }): void => {
        this.setState({

          Items: response.value
        });
      }, (error: any): void => {
        this.setState({
          status: 'Loading all items failed with error: ' + error,
          Items: []
        });
      });
  }

  private _getItemEntityType(): Promise<string> {
    return this.props.spHttpClient.get(
      `${this.props.siteUrl}/_api/web/lists/getById('${this.props.listName}')?$select=ListItementityTypeFullName`,
      SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.ListItementityTypeFullName
      }) as Promise<string>;
  }



  private _addListItem(): void {
    this.setState({
      status: ''
    });
    this._getItemEntityType().then(async spEntityType => {
      const request: any = {};
      request.body = JSON.stringify({
        Title: this.state.itemTitle,
        '@odata.type': spEntityType
      });

      const response = await this.props.spHttpClient.post(
        `${this.props.siteUrl}/_api/web/lists/getById('${this.props.listName}')/items`,
        SPHttpClient.configurations.v1, request
      )
      if (response.ok) {
        this.setState({
          status: 'Item Add Success'
        });
        this._getListItems();
      }
      else {
        this.setState({
          status: 'Error while creating the item: ' + response.statusMessage
        });
      }
    })
  }
  private _updateListItem(): void {
    this.setState({
      status: ''
    });
    this.props.spHttpClient.get(
      `${this.props.siteUrl}/_api/web/lists/getById('${this.props.listName}')/items(${this.state.itemId})`,
      SPHttpClient.configurations.v1).then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.value;
      })
      .then(async (listItem: IListItem) => {

        const request: any = {};
        request.headers = {
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': "*"
        };
        request.body = JSON.stringify({ 'Title': `${this.state.itemTitle}` });

        const response = await this.props.spHttpClient.post(
          `${this.props.siteUrl}/_api/web/lists/getById('${this.props.listName}')/items(${this.state.itemId})`,
          SPHttpClient.configurations.v1, request);

        if (response.ok) {
          this.setState({
            status: 'Item Update Success'
          });
          this._getListItems();
        }
        else {
          this.setState({
            status: 'Error while updating the item: ' + response.statusMessage
          });
        }
      })
  }


  private _deleteListItem(): void {

    this.props.spHttpClient.get(
      `${this.props.siteUrl}/_api/web/lists/getById('${this.props.listName}')/items(${this.state.itemId})`,
      SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.value;
      })
      .then(async (listItem: IListItem) => {
        const request: any = {};
        request.headers = {
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*'
        };
        request.body = JSON.stringify(listItem);

        const response = await this.props.spHttpClient.post(
          `${this.props.siteUrl}/_api/web/lists/getById('${this.props.listName}')/items(${this.state.itemId})`,
          SPHttpClient.configurations.v1,
          request);

        if (response.ok) {
          this.setState({
            status: 'Item Update Success'
          });
          this._getListItems();
        }
        else {
          this.setState({
            status: 'Error while updating the item: ' + response.statusMessage
          });
        }

      });

  }


}
