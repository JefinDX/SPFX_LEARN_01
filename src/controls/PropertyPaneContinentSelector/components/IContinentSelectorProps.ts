/* eslint-disable @typescript-eslint/no-explicit-any */
import { IDropdownOption } from '@fluentui/react';

export interface IContinentSelectorProps {
  label: string;
  onChangedReactive: (option: IDropdownOption, index?: number) => void;
  onChangedNonReactive: (targetProperty?: string, newValue?: any) => void;
  selectedKey: string | number;
  disabled: boolean;
  stateKey: string;
}