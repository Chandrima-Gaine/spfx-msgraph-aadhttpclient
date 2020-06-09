import * as React from 'react';
import styles from './MsGraphAadHttpClient.module.scss';
import { IMsGraphAadHttpClientProps } from './IMsGraphAadHttpClientProps';
import { IMsGraphAadHttpClientState } from './IMsGraphAadHttpClientState';
import { IUserItem } from './IUserItem';
import { escape } from '@microsoft/sp-lodash-subset';

import { AadHttpClient } from '@microsoft/sp-http';

import {
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode
} from 'office-ui-fabric-react';

// Configure the columns for the DetailsList component
let _usersListColumns = [
  {
    key: 'displayName',
    name: 'Display name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'mail',
    name: 'Mail',
    fieldName: 'mail',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'userPrincipalName',
    name: 'User Principal Name',
    fieldName: 'userPrincipalName',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
];

export default class MsGraphAadHttpClient extends React.Component<IMsGraphAadHttpClientProps, IMsGraphAadHttpClientState> {
  constructor(props: IMsGraphAadHttpClientProps, state: IMsGraphAadHttpClientState) {
    super(props);
    
    // Initialize the state of the component
    this.state = {
      users: []
    };
  }

  public render(): React.ReactElement<IMsGraphAadHttpClientProps> {
    return (
      <div className={ styles.msGraphAadHttpClient }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>SharePoint Framework</span>
              <p className={ styles.subTitle }>Consume Microsoft Graph API Using AadHttpClient</p>
              <p className={ styles.description }>Click the button to get User details...</p>
              
              <p className={ styles.form }>
                <PrimaryButton 
                    text='User Details' 
                    title='User Details' 
                    onClick={ this.getUserDetails } 
                  />
              </p>
              {
                (this.state.users != null && this.state.users.length > 0) ?
                  <p className={ styles.form }>
                  <DetailsList
                      items={ this.state.users }
                      columns={ _usersListColumns }
                      setKey='set'
                      checkboxVisibility={ CheckboxVisibility.hidden }
                      selectionMode={ SelectionMode.none }
                      layoutMode={ DetailsListLayoutMode.fixedColumns }
                      compact={ true }
                  />
                </p>
                : null
              }

            </div>
          </div>
        </div>
      </div>
    );
  }

  @autobind
  private getUserDetails(): void {
    const aadClient: AadHttpClient = new AadHttpClient(
      this.props.context.serviceScope,
      "https://graph.microsoft.com"
    );

    // Get users with givenName, surname, or displayName
    aadClient
      .get(
        `https://graph.microsoft.com/v1.0/users?$select=displayName,mail,userPrincipalName`,
        AadHttpClient.configurations.v1
      )
      .then(response => {
        return response.json();
      })
      .then(json => {
        // Prepare the output array
        var users: Array<IUserItem> = new Array<IUserItem>();

        // Log the result in the console for testing purposes
        console.log(json);

        // Map the JSON response to the output array
        json.value.map((item: any) => {
          users.push( { 
            displayName: item.displayName,
            mail: item.mail,
            userPrincipalName: item.userPrincipalName,
          });
        });

        // Update the component state accordingly to the result
        this.setState(
          {
            users: users,
          }
        );
      })
      .catch(error => {
        console.error(error);
      });
  }
}
