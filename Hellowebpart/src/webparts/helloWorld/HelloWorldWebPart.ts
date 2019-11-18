import {
  Version,
  DisplayMode, //Sayfanın edit yada display olup olmadığımı control eder
  Environment,
  EnvironmentType, //Localde mi çalışıyoruz Online da mı kontrolü,
  Log, //Console logları yazar
} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider,

} from '@microsoft/sp-webpart-base';

///https://sharepoint.github.io/sp-dev-fx-property-controls/
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';


export interface IHelloWorldWebPartProps {
  description: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
  mycontentVisited:number,
  myValidateContent:string,

}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    const pageMode: string = this.displayMode === DisplayMode.Edit ? 'You are in Edit Mode'
      : 'You are in Display Mode';

    const environmentType: string = Environment.type === EnvironmentType.Local ? 'You are in local'
      : 'You are in Server';

    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "message");
    setTimeout(() => {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);

      this.domElement.innerHTML = `
      <div class="${ styles.helloWorld}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Cunpmstomize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <p class="${ styles.description}"><strong>My Visited Count : </strong>${this.properties.mycontentVisited}</p>


              <p class="${ styles.description}">${escape(this.properties.test)}</p>
              <p class="${ styles.description}">${escape(this.properties.myValidateContent)}</p>
              <br />
              <p class="${ styles.subTitle}"><strong> Page Mode : </strong>${escape(pageMode)}</p>
              <p class="${ styles.subTitle}"><strong> Environment: </strong>${escape(environmentType)}</p>
              <a href="#" class="${ styles.button}">
                <span class="${ styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;

      this.domElement.getElementsByClassName(`${styles.button}`)[0]
        .addEventListener('click', (event: any) => {
          event.preventDefault();
          alert('Welcome to Sharepoint Framework');
        });
        Log.info('Helloworld', 'message',this.context.serviceScope);
        Log.warn('Helloworld', 'WARNİNG message',this.context.serviceScope);
        Log.error('Helloworld', new Error('Error Message'),this.context.serviceScope);
        Log.verbose('Helloworld', 'VERBOSE Mesaage',this.context.serviceScope);
      }, 3000)
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
                PropertyPaneTextField('test', {
                  label: 'Multi Line Text Field',
                  multiline: true,

                }),

                PropertyPaneCheckbox('test1', {
                  text: 'CheckBox'
                }),

                PropertyPaneDropdown('test2', {
                  label: 'Dropdown',
                  options: [
                    { key: '1', text: 'One' },
                    { key: '2', text: 'Two' },
                    { key: '3', text: 'Three' },
                    { key: '4', text: 'Four' },
                  ]
                }),

                PropertyPaneToggle('test3', {
                  label: 'Toggle',
                  onText: 'On',
                  offText: 'Off'
                }),

                PropertyPaneSlider('mycontentVisited', {
                  label:'We visited Count',
                  min:1,
                  max:10,
                  showValue:true
                }),


              /*   PropertyPaneTextField('myValidateContent' ,{
                  label:'Validate Text Field',
                  onGetErrorMessage:this.ValidateContinents.bind(this)
                }) */


              ]
            }
          ]
        }
      ]
    };

  }
/*   private ValidateContinents(textboxValue : string) :string{
    const ValidateCOntinentOptions: string [] =['africa', 'antartica', 'asia', 'north america', 'south america'];
    const inputToValidate:string = textboxValue.toLowerCase();

    return (ValidateCOntinentOptions.indexOf(inputToValidate)===-1)
    ? 'Invalid continent entry; valid options are "Africa"' : ''
  }
 */
  private onContinentSelectionChange(propertPath : string, newValue: any): void {
    const oldValue : any =this.properties[propertPath];
    this.properties[propertPath]=newValue;
    this.render();
  }
}
