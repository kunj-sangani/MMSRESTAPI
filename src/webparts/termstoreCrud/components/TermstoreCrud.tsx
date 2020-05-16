import * as React from 'react';
import styles from './TermstoreCrud.module.scss';
import { ITermstoreCrudProps } from './ITermstoreCrudProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as msal from "@azure/msal-browser";
import { IAllGroups, IAllSets, IAllTerms } from "./TermstoreInterfaces";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
} from 'office-ui-fabric-react/lib/DetailsList';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { PrimaryButton, IRenderFunction, IconButton } from 'office-ui-fabric-react';
import GroupPanel from "./GroupPanel";

const msalConfig = {
  auth: {
    clientId: '6bbe73c1-3bbd-41ce-97c5-0bede0bf2554',
    redirectUri: 'https://testinglala.sharepoint.com/_layouts/15/workbench.aspx'
  }
};

export interface ITermstoreCrudState {
  groups?: IAllGroups[];
  sets?: IAllSets[];
  terms?: IAllTerms[];
  displaygroups?: boolean;
  displayterms?: boolean;
  selectedGroup?: string;
  selectedset?: string;
  isOpenGroup?: boolean;
  selectedpanelGroup?: IAllGroups;
}


const msalInstance = new msal.PublicClientApplication(msalConfig);
export default class TermstoreCrud extends React.Component<ITermstoreCrudProps, ITermstoreCrudState> {
  private _selection: Selection;
  constructor(props: ITermstoreCrudProps, state: ITermstoreCrudState) {
    super(props);
    this._selection = new Selection();
    this.state = {
      groups: [],
      sets: [],
      displaygroups: true,
      displayterms: false,
      isOpenGroup: false
    };
  }

  public allGroupColumns: IColumn[] = [
    {
      key: 'name',
      name: 'Group Name',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IAllGroups) => {
        return <div>{item.name}</div>;
      },
    },
    {
      key: 'description',
      name: 'Description',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IAllGroups) => {
        return <div>{item.description}</div>;
      },
    },
    {
      key: 'createdDateTime',
      name: 'created Date Time',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IAllGroups) => {
        return <div>{new Date(item.createdDateTime).toDateString()}</div>;
      },
    },
    {
      key: 'lastModifiedDateTime',
      name: 'lastModified Date Time',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IAllGroups) => {
        return <div>{new Date(item.lastModifiedDateTime).toDateString()}</div>;
      },
    }
  ];
  public allSetsColumns: IColumn[] = [
    {
      key: 'name',
      name: 'Sets Name',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IAllSets) => {
        return (item.localizedNames.map((val) => {
          return (<div>{val.name} - {val.languageTag}</div>);
        }));
      },
    },
    {
      key: 'description',
      name: 'Description',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IAllSets) => {
        return <div>{item.description}</div>;
      },
    },
    {
      key: 'createdDateTime',
      name: 'created Date Time',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IAllSets) => {
        return <div>{new Date(item.createdDateTime).toDateString()}</div>;
      },
    },
    {
      key: 'childrenCount',
      name: 'children Count',
      minWidth: 70,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IAllSets) => {
        return <div>{item.childrenCount}</div>;
      },
    }
  ];
  public alltermsColumns: IColumn[] = [
    {
      key: 'name',
      name: 'Term Name',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IAllTerms) => {
        return (item.labels.map((val) => {
          return (<div>{val.name} - {val.languageTag}</div>);
        }));
      },
    },
    {
      key: 'description',
      name: 'Description',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IAllTerms) => {
        return (item.descriptions.map((val) => {
          return (<div>{val.description} - {val.languageTag}</div>);
        }));
      },
    },
    {
      key: 'createdDateTime',
      name: 'created Date Time',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IAllTerms) => {
        return <div>{new Date(item.createdDateTime).toDateString()}</div>;
      },
    },
    {
      key: 'childrenCount',
      name: 'children Count',
      minWidth: 70,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IAllTerms) => {
        return <div>{item.childrenCount}</div>;
      },
    }
  ];

  public silentRequest = {
    scopes: ["https://testinglala.sharepoint.com/TermStore.Read.All", "https://testinglala.sharepoint.com/TermStore.ReadWrite.All"],
    loginHint: this.props.webpartContext.pageContext.user.email
  };

  public getTermStoreData = (tokenResponse, endpoint: string): Promise<any> => {
    let headers = new Headers();
    let bearer = "Bearer " + tokenResponse.accessToken;
    headers.append("Authorization", bearer);
    let options = {
      method: "GET",
      headers: headers
    };
    let graphEndpoint = endpoint;

    return new Promise<any>((resolve, reject) => {
      fetch(graphEndpoint, options)
        .then(resp => {
          resp.json().then((groups) => {
            resolve(groups);
          }).catch((error) => {
            reject(error);
          });
        });
    });
  }

  public updateTermStoreData = (tokenResponse, endpoint: string, termstorebody): Promise<any> => {
    let headers = new Headers();
    let bearer = "Bearer " + tokenResponse.accessToken;
    headers.append("Authorization", bearer);
    headers.append("Content-Type", "application/json");
    let options = {
      method: "PATCH",
      headers: headers,
      body: JSON.stringify(termstorebody)
    };
    let graphEndpoint = endpoint;

    return new Promise<any>((resolve, reject) => {
      fetch(graphEndpoint, options)
        .then(resp => {
          if (resp.status === 200) {
            resolve(true);
          } else {
            resolve(false);
          }
        }).catch(err => {
          reject(err);
        });
    });
  }

  public addTermStoreData = (tokenResponse, endpoint: string, termstorebody): Promise<any> => {
    let headers = new Headers();
    let bearer = "Bearer " + tokenResponse.accessToken;
    headers.append("Authorization", bearer);
    headers.append("Content-Type", "application/json");
    let options = {
      method: "POST",
      headers: headers,
      body: JSON.stringify(termstorebody)
    };
    let graphEndpoint = endpoint;

    return new Promise<any>((resolve, reject) => {
      fetch(graphEndpoint, options)
        .then(resp => {
          if (resp.status === 200) {
            resolve(true);
          } else {
            resolve(true);
          }
        }).catch(err => {
          reject(err);
        });
    });
  }

  public deleteTermStoreData = (tokenResponse, endpoint: string): Promise<any> => {
    let headers = new Headers();
    let bearer = "Bearer " + tokenResponse.accessToken;
    headers.append("Authorization", bearer);
    let options = {
      method: "DELETE",
      headers: headers
    };
    let graphEndpoint = endpoint;

    return new Promise<any>((resolve, reject) => {
      fetch(graphEndpoint, options)
        .then(resp => {
          if (resp.status === 200) {
            resolve(true);
          } else {
            resolve(true);
          }
        }).catch(err => {
          reject(err);
        });
    });
  }

  public componentDidMount() {
    msalInstance.ssoSilent(this.silentRequest).then((response) => {
      console.log(response);
      this._getGroups();
    }).catch((error) => {

    });
  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }

  private _getGroups = (): void => {
    msalInstance.acquireTokenSilent(this.silentRequest).then((val) => {
      console.log(val);
      this.getTermStoreData(val, "https://testinglala.sharepoint.com/_api/v2.1/termStore/groups").then((groups) => {
        this.setState({
          groups: groups.value
        });
      }).catch((error) => {
        console.log(error);
      });
    }).catch((err) => {
      console.log(err);
    });
  }

  private _onItemInvoked = (item: any): void => {
    //alert(`Item invoked: ${item.id}`);
    msalInstance.acquireTokenSilent(this.silentRequest).then((val) => {
      this.getTermStoreData(val, `https://testinglala.sharepoint.com/_api/v2.1/termStore/groups/${item.id}/sets`).then((sets) => {
        this.setState({
          sets: sets.value,
          displaygroups: false,
          selectedGroup: item.name
        });
      }).catch((error) => {
        console.log(error);
      });
    }).catch((err) => {
      console.log(err);
    });
  }

  private _onItemInvokedsets = (item: any): void => {
    msalInstance.acquireTokenSilent(this.silentRequest).then((val) => {
      this.getTermStoreData(val, `https://testinglala.sharepoint.com/_api/v2.1/termStore/groups/${item.groupId}/sets/${item.id}/terms`).then((terms) => {
        this.setState({
          terms: terms.value,
          displayterms: true,
          selectedset: item.localizedNames[0].name
        });
      }).catch((error) => {
        console.log(error);
      });
    }).catch((err) => {
      console.log(err);
    });
  }

  private _onItemInvokedterms = (item: any): void => {
  }

  private _onEditGroupItem = (item: any, termstorebody: any): void => {
    //alert(`Item invoked: ${item.id}`);
    msalInstance.acquireTokenSilent(this.silentRequest).then((val) => {
      this.updateTermStoreData(val, `https://testinglala.sharepoint.com/_api/v2.1/termStore/groups/${item.id}`, termstorebody).then((status) => {
        if (status) {
          this.dismissPanelGroup();
          this._getGroups();
        }
      }).catch((error) => {
        console.log(error);
      });
    }).catch((err) => {
      console.log(err);
    });
  }

  private _onAddGroupItem = (termstorebody: any): void => {
    msalInstance.acquireTokenSilent(this.silentRequest).then((val) => {
      this.addTermStoreData(val, `https://testinglala.sharepoint.com/_api/v2.1/termStore/groups`, termstorebody).then((status) => {
        if (status) {
          this.dismissPanelGroup();
          this._getGroups();
        }
      }).catch((error) => {
        console.log(error);
      });
    }).catch((err) => {
      console.log(err);
    });
  }

  private _onDeleteGroupItem = (termstorebody: any): void => {
    msalInstance.acquireTokenSilent(this.silentRequest).then((val) => {
      this.deleteTermStoreData(val, `https://testinglala.sharepoint.com/_api/v2.1/termStore/groups/${termstorebody.id}`).then((status) => {
        if (status) {
          this.dismissPanelGroup();
          this._getGroups();
        }
      }).catch((error) => {
        console.log(error);
      });
    }).catch((err) => {
      console.log(err);
    });
  }


  private goback = () => {
    this.setState({
      displaygroups: true
    });
  }

  private gobacktosets = () => {
    this.setState({
      displayterms: false
    });
  }

  private editGroup = () => {
    if (this._selection.getSelectedCount() > 0) {
      let selectedGroup = this._selection.getSelection()[0] as IAllGroups;
      this.setState({
        isOpenGroup: true,
        selectedpanelGroup: selectedGroup
      });
    }
  }

  private addGroup = () => {
    this.setState({
      isOpenGroup: true,
      selectedpanelGroup: null
    });
  }

  private deleteGroup = () => {
    if (this._selection.getSelectedCount() > 0) {
      let selectedGroup = this._selection.getSelection()[0] as IAllGroups;
      this._onDeleteGroupItem(selectedGroup);
    }
  }

  private dismissPanelGroup = () => {
    this.setState({
      isOpenGroup: false
    });
  }

  public render(): React.ReactElement<ITermstoreCrudProps> {
    return (
      <div className={styles.termstoreCrud} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column} style={{ display: this.state.displaygroups ? "block" : "none" }}>
              <h1>All Groups</h1>
              <h3>Double click on the Group to fetch term sets</h3>
              <PrimaryButton iconProps={{ iconName: "Add" }} text="Add Group" onClick={this.addGroup} style={{ marginRight: 10 }} />
              <PrimaryButton iconProps={{ iconName: "Edit" }} text="Edit Group" onClick={this.editGroup} style={{ marginRight: 10 }} />
              <PrimaryButton iconProps={{ iconName: "Delete" }} text="Delete Group" onClick={this.deleteGroup} />
              <DetailsList
                items={this.state.groups ? this.state.groups : []}
                compact={false}
                columns={this.allGroupColumns}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvoked}
                selectionMode={SelectionMode.single}
                selection={this._selection}
              />
            </div>
            <div className={styles.column} style={{ display: !this.state.displaygroups && !this.state.displayterms ? "block" : "none" }}>
              <h1>All term sets for {this.state.selectedGroup}</h1>
              <h3 onClick={this.goback}><Icon iconName="ChromeBack" /><span> Go Back</span></h3>
              <PrimaryButton iconProps={{ iconName: "Add" }} text="Add Set" onClick={this.addGroup} style={{ marginRight: 10 }} />
              <PrimaryButton iconProps={{ iconName: "Edit" }} text="Edit Set" onClick={this.editGroup} style={{ marginRight: 10 }} />
              <PrimaryButton iconProps={{ iconName: "Delete" }} text="Delete Set" onClick={this.deleteGroup} />
              <DetailsList
                items={this.state.sets ? this.state.sets : []}
                compact={false}
                columns={this.allSetsColumns}
                selectionMode={SelectionMode.single}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvokedsets}

              />
            </div>
            <div className={styles.column} style={{ display: this.state.displayterms ? "block" : "none" }}>
              <h1>All term for {this.state.selectedset}</h1>
              <h3 onClick={this.gobacktosets}><Icon iconName="ChromeBack" /><span> Go Back</span></h3>
              <DetailsList
                items={this.state.terms ? this.state.terms : []}
                compact={false}
                columns={this.alltermsColumns}
                selectionMode={SelectionMode.single}
                getKey={this._getKey}
                setKey="none"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                onItemInvoked={this._onItemInvokedterms}

              />
            </div>
          </div>
          <GroupPanel isOpenGroup={this.state.isOpenGroup} dismissPanelGroup={this.dismissPanelGroup}
            selectedpanelGroup={this.state.selectedpanelGroup} onUpdate={this._onEditGroupItem} onAdd={this._onAddGroupItem}></GroupPanel>
        </div>
      </div>
    );
  }
}
