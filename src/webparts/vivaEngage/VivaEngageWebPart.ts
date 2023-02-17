/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';


import styles from './VivaEngageWebPart.module.scss';
import * as strings from 'VivaEngageWebPartStrings';

import { AadTokenProvider, HttpClient, HttpClientResponse} from '@microsoft/sp-http';

export interface IVivaEngageWebPartProps {
  description: string;
}

export default class VivaEngageWebPart extends BaseClientSideWebPart<IVivaEngageWebPartProps> {

  private vivaEngageToken: string = '';
  private vivaEngagePosts: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.vivaEngage} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div>
        <h3>Welcome to Viva Engage on SharePoint!</h3>
        <p>
        This is a sample example from where you will be able to connect to Viva Engage using SPFx and retrieve the latest updates for the current user.
        </p>
        <h4>Last Viva Engage updates for you:</h4>
          <ul class="${styles.links}">
            ${this.vivaEngagePosts}
          </ul>
      </div>
    </section>`;
  }

  protected async onInit(): Promise<void> {
    await this.getViVaEngageToken();
    await this.getPosts();
  }

  private async getViVaEngageToken(){
    const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();    
    await tokenProvider.getToken("https://api.yammer.com").then(async token => {
      this.vivaEngageToken = token;
    });
  }

  private async getPosts() {

    let tempData: any;
    if(this.vivaEngagePosts!=='') return;

    await this.context.httpClient.get(`https://api.yammer.com/api/v1/messages.json?limit=30&threaded=true`,
        HttpClient.configurations.v1,
        {
            headers: {
            "Authorization": `Bearer ${this.vivaEngageToken}`,
            'Content-type': 'application/json'
            }
        }
        ).then((response: HttpClientResponse) => {
            return response.json();
        }).then((data: object) => {            
            tempData = data;
            tempData.messages.forEach((message: any) => {   
              this.vivaEngagePosts += `<li><a href="${message.web_url}" target="_blank">${message.body.plain}</a></li>`;                      
            }); 
        }, (err: any): void => {
            console.log(err);
        });
  }





  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
