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
// import { escape } from '@microsoft/sp-lodash-subset';
import { escape, update } from '@microsoft/sp-lodash-subset';
import { DisplayMode, EnvironmentType, Environment, Log } from '@microsoft/sp-core-library';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import {
  PropertyPaneContinentSelector,
  IPropertyPaneContinentSelectorProps
} from '../../controls/PropertyPaneContinentSelector';


export interface IHelloWorldWebPartProps {
  description: string;
  myContinent1: string;
  myContinent2: string;
  numContinentsVisited: number;
  customField?: string;
}

export interface IHelloWorldWebPartState {
  descriptionState: string;
}

export interface IHelloPropertyPaneWebPartProps {
  description: string;
  myContinent: string;
  numContinentsVisited: number;
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
        <div>Continent 1 where I reside: <strong>${escape(this.properties.myContinent1)}</strong></div>
        <div>Continent 2 where I reside: <strong>${escape(this.properties.myContinent2)}</strong></div>
        <div>Number of continents I've visited: <strong>${this.properties.numContinentsVisited}</strong></div>
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
            description: strings.PropertyPaneDescription
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
                  key: 'myContinentKey',
                  label: 'Continent 2 where I currently reside',
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
        }
      ]
    };
  }

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
