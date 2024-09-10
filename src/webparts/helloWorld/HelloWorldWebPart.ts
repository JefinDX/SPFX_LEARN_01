import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
/**
 * PropertyPaneCustomField is discontinued from spfx v1.19.0
 */
//import { PropertyPaneCustomField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape, update } from '@microsoft/sp-lodash-subset';
import { DisplayMode, EnvironmentType, Environment, Log } from '@microsoft/sp-core-library';

import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import {
  IPropertyFieldGroupOrPerson,
  PropertyFieldPeoplePicker,
  PrincipalType
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import {
  PropertyPaneContinentSelector,
  PropertyPaneContinentSelectorNonReactive,
  IPropertyPaneContinentSelectorProps
} from '../../controls/PropertyPaneContinentSelector';

export interface IHelloWorldWebPartProps {
  description: string;
  myContinent1: string;
  myContinent2: string;
  myContinent3: string;
  numContinentsVisited: number;
  customField?: string;
  lists: string;
  people: IPropertyFieldGroupOrPerson[];
  expansionOptions: any[]; // eslint-disable-line @typescript-eslint/no-explicit-any
}

export interface IHelloWorldWebPartState {
  descriptionState: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  public render(): void {
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "message");
    setTimeout(() => {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this._renderWebpartDOM();
    }, 2000)

    /**
     * SPFx logging example
     */
    Log.info('HelloWorld', 'message', this.context.serviceScope);
    Log.warn('HelloWorld', 'WARNING message', this.context.serviceScope);
    Log.error('HelloWorld', new Error('Error message'), this.context.serviceScope);
    Log.verbose('HelloWorld', 'VERBOSE message', this.context.serviceScope);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _renderWebpartDOM(): void {
    const siteTitle: string = this.context.pageContext.web.title;
    const pageMode: string = (this.displayMode === DisplayMode.Edit)
      ? 'You are in edit mode'
      : 'You are in read mode';
    const environmentType: string = (Environment.type === EnvironmentType.ClassicSharePoint)
      ? 'You are running in a classic page'
      : 'You are running in a modern page';
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <h3>WebPart Property Pane Values</h3>

        <div>Web part property value description: <strong>${escape(this.properties.description)}</strong></div>
        ${this.properties.myContinent1 ?
        `<div>Continent 1 where I reside: <strong>${escape(this.properties.myContinent1)}</strong></div>` : ''}
        ${this.properties.myContinent2 ?
        `<div>Continent 2 where I reside: <strong>${escape(this.properties.myContinent2)}</strong></div>` : ''}
          ${this.properties.myContinent3 ?
        `<div>Continent 3 where I reside: <strong>${escape(this.properties.myContinent3)}</strong></div>` : ''}
            ${this.properties.numContinentsVisited ?
        `<div>Number of continents I've visited: <strong>${this.properties.numContinentsVisited}</strong></div>` : ''}

        <div>List selected: <strong>${this.properties.lists}</strong></div>
        <div class="selectedPeople"></div>
        <div class="expansionOptions"></div>

        <h3>SharePoint Details</h3>
        <div>Site title: <strong>${escape(siteTitle)}</strong></div>
        <div>Page mode: <strong>${escape(pageMode)}</strong></div>
        <div>Environment: <strong>${escape(environmentType)}</strong></div>
      </div>
      <button type="button">Show Loading (status renderer)</button>
      </div>
    </section>`;

    this.domElement.getElementsByTagName("button")[0]
      .addEventListener('click', (event: MouseEvent) => {
        event.preventDefault();
        this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Welcome to the SharePoint Framework!");
        setTimeout(() => {
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this._renderWebpartDOM();
        }, 5000);
      });

    if (this.properties.people && this.properties.people.length > 0) {
      let peopleList: string = '';
      this.properties.people.forEach((person) => {
        peopleList = peopleList + `<li>${person.fullName} (${person.email})</li>`;
      });
      this.domElement.getElementsByClassName('selectedPeople')[0].innerHTML = `<ul>${peopleList}</ul>`;
    }

    if (this.properties.expansionOptions && this.properties.expansionOptions.length > 0) {
      let expansionOptions: string = '';
      this.properties.expansionOptions.forEach((option) => {
        expansionOptions = expansionOptions + `<li>${option.Region}: ${option.Comment} </li>`;
      });
      if (expansionOptions.length > 0) {
        this.domElement.getElementsByClassName('expansionOptions')[0].innerHTML = `<ul>${expansionOptions}</ul>`;
      }
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
            // description: strings.PropertyPaneDescription
            description: 'Built-in and Custom Property Pane Controls'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('myContinent1', {
                  label: 'Continent 1 where I currently reside',
                  onGetErrorMessage: this._validateContinents.bind(this)
                }),
                /**
                 * PropertyPaneCustomField is discontinued from spfx v1.19.0
                 */
                // PropertyPaneCustomField('customField', { 
                //   onRender: this._customFieldRender.bind(this)
                //  } ),
                new PropertyPaneContinentSelector('myContinent2', <IPropertyPaneContinentSelectorProps>{
                  key: 'myContinent2Key',
                  label: 'Continent 2 where I currently reside',
                  disabled: false,
                  selectedKey: this.properties.myContinent2,
                  onPropertyChange: this.onContinentSelectionChange.bind(this),
                }),
                new PropertyPaneContinentSelectorNonReactive('myContinent3', <IPropertyPaneContinentSelectorProps>{
                  key: 'myContinent3Key',
                  label: 'Continent 3 where I currently reside',
                  disabled: false,
                  selectedKey: this.properties.myContinent2,
                  onPropertyChange: this.onContinentSelectionChange.bind(this),
                }),
                PropertyPaneSlider('numContinentsVisited', {
                  label: 'Number of continents I\'ve visited',
                  min: 1, max: 7, showValue: true,
                })
              ]
            }
          ]
        },
        {
          header: {
            description: 'PnP Control Fields'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldPeoplePicker('people', {
                  label: 'People Picker',
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context as any, // eslint-disable-line @typescript-eslint/no-explicit-any
                  properties: this.properties,
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
                        { key: 'East', text: 'East' },
                        { key: 'West', text: 'West' },
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
                })
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * PropertyPaneCustomField is discontinued from spfx v1.19.0
   */
  // private _customFieldRender(elem: HTMLElement): void {
  //   elem.innerHTML = '<div><h1>This is a custom field.</h1></div>';
  // }

  private _validateContinents(textboxValue: string): string {
    const validContinentOptions: string[] = ['africa', 'antarctica', 'asia', 'australia', 'europe', 'north america', 'south america'];
    const inputToValidate: string = textboxValue.toLowerCase();
    return (validContinentOptions.indexOf(inputToValidate) === -1)
      ? 'Invalid continent entry; valid options are "Africa", "Antarctica", "Asia", "Australia", "Europe", "North America", and "South America"'
      : '';
  }

  /* eslint-disable @typescript-eslint/no-explicit-any */
  private onContinentSelectionChange(propertyPath: string, newValue: any): void {
    update(this.properties, propertyPath, (): any => { return newValue });
    this.render();
  }
  /* eslint-enable @typescript-eslint/no-explicit-any */
}
