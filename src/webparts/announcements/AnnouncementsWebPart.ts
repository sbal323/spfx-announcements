import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AnnouncementsWebPart.module.scss';
import * as strings from 'AnnouncementsWebPartStrings';

import "../../../../sharepoint-angular/dist/sharepoint-angular/main";
import "../../../../sharepoint-angular/dist/sharepoint-angular/polyfills";
import "../../../../sharepoint-angular/dist/sharepoint-angular/runtime";
require("../../../../sharepoint-angular/dist/sharepoint-angular/scripts");
require("../../../../sharepoint-angular/dist/sharepoint-angular/styles.css");

export interface IAnnouncementsWebPartProps {
  description: string;
}

export default class AnnouncementsWebPart extends BaseClientSideWebPart<IAnnouncementsWebPartProps> {

  public render(): void {
    // this.domElement.innerHTML = `
    //   <div class="${ styles.announcements }">
    //     <div class="${ styles.container }">
    //       <div class="${ styles.row }">
    //         <div class="${ styles.column }">
    //           <span class="${ styles.title }">Welcome to SharePoint!</span>
    //           <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
    //           <p class="${ styles.description }">${escape(this.properties.description)}</p>
    //           <a href="https://aka.ms/spfx" class="${ styles.button }">
    //             <span class="${ styles.label }">Learn more</span>
    //           </a>
    //         </div>
    //       </div>
    //     </div>
    //   </div>`;
    window['_spPageContextInfo'] = this.context.pageContext.legacyPageContext;
    this.domElement.innerHTML = `<announcements-element></announcements-element>`;
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
