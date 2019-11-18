//https://github.com/SharePoint/sp-dev-training-spfx-webpart-proppane/tree/master/Demos/03-pnp-controls

import { Version } from '@microsoft/sp-core-library';
import * as React from 'react';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';




import {
  IPropertyFieldGroupOrPerson,
  PropertyFieldPeoplePicker,
  PrincipalType
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle} from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldDateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldOrder } from '@pnp/spfx-property-controls/lib/PropertyFieldOrder';
import {orderedItem} from '../../controls/PnpOrderedField/orderedField'



import styles from './PnpCustomFieldWebPart.module.scss';
import * as strings from 'PnpCustomFieldWebPartStrings';

export interface IPnpCustomFieldWebPartProps {
  description: string;
  people: IPropertyFieldGroupOrPerson[];
  expansionOptions: any[];
  color: string;
  datetime: IDateTimeFieldValue;
  lists: string | string[];
  orderedItems: Array<any>;
}

export default class PnpCustomFieldWebPart extends BaseClientSideWebPart<IPnpCustomFieldWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.pnpCustomField }" styles="background-color:${ this.properties.color}" >
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
            <div class="selectedPeople"></div>
            <div class="expansionOptions"></div>
            <div class="datetime"><p>${this.properties.datetime ? this.properties.datetime['displayValue'] :
           '' }</p></div>
            ${this.properties.color ? '<div style="width:100px;height:100px;background-color:'+this.properties.color +'"></div>' : null}
          </div>
        </div>
      </div>`;

      if (this.properties.people && this.properties.people.length > 0) {
        let peopleList: string = '';
        this.properties.people.forEach((person) => {
          peopleList = peopleList + `<li>${ person.fullName } (${ person.email })</li>`;
        });

        this.domElement.getElementsByClassName('selectedPeople')[0].innerHTML = `<ul>${ peopleList }</ul>`;
      }

     if (this.properties.expansionOptions && this.properties.expansionOptions.length > 0) {
        let expansionOptions: string  = '';
        this.properties.expansionOptions.forEach((option) => {
          expansionOptions = expansionOptions + `<li>${ option['Region'] }: ${ option['Comment'] } </li>`;
        });
        if (expansionOptions.length > 0) {
          this.domElement.getElementsByClassName('expansionOptions')[0].innerHTML = `<ul>${ expansionOptions }</ul>`;
        }
      }


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
                }),

                PropertyFieldPeoplePicker('people', {
                  label: 'Property Pane Field People Picker PnP Reusable Control',
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                }),

                PropertyFieldCollectionData('expansionOptions', {
                  key: 'collectionData',
                  label: 'Possible expansion options',
                  panelHeader: 'Possible expansion options',
                  manageBtnLabel: 'Manage expansion options',
                  value: this.properties.expansionOptions,
                  fields: [
                    {
                      id: 'Region',
                      title: 'Region',
                      required: true,
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        { key: 'Northeast', text: 'Northeast' },
                        { key: 'Northwest', text: 'Northwest' },
                        { key: 'Southeast', text: 'Southeast' },
                        { key: 'Southwest', text: 'Southwest' },
                        { key: 'North', text: 'North' },
                        { key: 'South', text: 'South' }
                      ]
                    },
                    {
                      id: 'Comment',
                      title: 'Comment',
                      type: CustomCollectionFieldType.string
                    }
                  ]
                }),


                PropertyFieldColorPicker('color', {
                  label: 'Color',
                  selectedColor: this.properties.color,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),

                PropertyFieldDateTimePicker('datetime', {
                  label: 'Select the date and time',
                  initialDate: this.properties.datetime,
                  dateConvention: DateConvention.DateTime,
                  timeConvention: TimeConvention.Hours12,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'dateTimeFieldId',
                  showLabels: false
                }),

                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  multiSelect:true,
                  baseTemplate:100 ,// only Get Custom List
                  /***
                   * 101 Document Library
                   * 102 Survey
                   * 103 Links
                   * 104 Announcements
                   * 105 Contacts
                   * 106 Calendar
                   * 107 Tasks
                   * 108 Discussion Board
                   * 109 Picture Library
                   * 110 DataSources
                   * 115 Form Library
                   * 1100 Issues Tracking
                   */
                  listsToExclude:["IsListesi","6ea993ca-724c-4304-b293-d8b1f74b21a4"]  // Not Show this Lists
                }),

                PropertyFieldOrder("orderedItems", {
                  key: "orderedItems",
                  label: "Ordered Items",
                  items: this.properties.orderedItems,
                  textProperty: "text",
                  //removeArrows: true,
                  //disableDragAndDrop: true,
                  onRenderItem: orderedItem,
                  //maxHeight: 90,
                  //disabled: true,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged


                })
              ]
            }
          ]
        }
      ]
    };
  }
   public onPropertyPaneFieldChanged() : void{


  }

}
