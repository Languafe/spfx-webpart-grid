import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GridWebPart.module.scss';
import * as strings from 'GridWebPartStrings';

import * as moment from 'moment';
import 'moment/locale/nb';

export interface IGridWebPartProps {
  description: string;
  title: string;
  numRows: number;
  numCols: number;
}

export default class GridWebPart extends BaseClientSideWebPart<IGridWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.grid }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">${this.properties.title}</span>
              <p class="clock">${moment().format('DD.MM.YYYY HH:mm:ss')}</p>
            </div>
          </div>
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              ${this.renderGrid()}
            </div>
          </div>
        </div>
      </div>`;
    
      this.toggleClock();
  }

  public toggleClock(): void {
    let container: Element = this.domElement.querySelector('.clock');

    setInterval(() => {
      container.innerHTML = moment().format('DD.MM.YYYY HH:mm:ss');
    }, 1000);    
  }

  public renderGrid(): string {

    let numRows: number = this.properties.numRows;
    let numCols: number = this.properties.numCols;

    let html = `<table class="${styles.table}">`;

    for (let i: number = 0; i < numRows; i++) {
      html += `<tr>`;
      for (let j: number = 0; j < numCols; j++) {
        html += `<td>${i+1},${j+1}</td>`;
      }
      html += `</tr>`;
    }

    html += `</table>`;
    
    return html;
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('numRows', {
                  label: strings.NumRowsFieldLabel,
                  maxLength: 2
                }),
                PropertyPaneTextField('numCols', {
                  label: strings.NumColsFieldLabel,
                  maxLength: 2
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
