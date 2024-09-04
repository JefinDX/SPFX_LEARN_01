import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
    IPropertyPaneField,
    PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
// import { IDropdownOption } from '@fluentui/react';
import { IContinentSelectorProps } from './components/IContinentSelectorProps';
import ContinentSelectorNonReactive from './components/ContinentSelectorNonReactive';
import {
    IPropertyPaneContinentSelectorProps,
    IPropertyPaneContinentSelectorInternalProps,
} from '.';

export class PropertyPaneContinentSelectorNonReactive implements IPropertyPaneField<IPropertyPaneContinentSelectorProps> {
    public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
    public properties: IPropertyPaneContinentSelectorInternalProps;
    private element: HTMLElement;

    constructor(public targetProperty: string, properties: IPropertyPaneContinentSelectorProps) {
        // debugger;
        this.properties = {
            key: properties.label,
            label: properties.label,
            disabled: properties.disabled,
            selectedKey: properties.selectedKey,
            onPropertyChange: properties.onPropertyChange,
            onRender: this.onRender.bind(this),
            onDispose: this.onDispose.bind(this)
        };
    }

    public render(): void {
        if (!this.element) {
            return;
        }
    }

    /* eslint-disable @typescript-eslint/no-explicit-any */
    private onRender(element: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        /* eslint-enable @typescript-eslint/no-explicit-any */
        if (!this.element) {
            this.element = element;
        }

        const reactElement: React.ReactElement<IContinentSelectorProps> = React.createElement(ContinentSelectorNonReactive, <IContinentSelectorProps>{
            label: this.properties.label,
            // onChangedReactive: this.onChanged.bind(this),
            onChangedNonReactive: changeCallback,
            selectedKey: this.properties.selectedKey,
            disabled: this.properties.disabled,
            stateKey: new Date().toString() // hack to allow for externally triggered re-rendering
        });
        ReactDom.render(reactElement, element);
    }

    private onDispose(element: HTMLElement): void {
        ReactDom.unmountComponentAtNode(element);
    }

    // private onChanged(option: IDropdownOption, index?: number): void {
    //     this.properties.onPropertyChange(this.targetProperty, option.key);
    // }
}