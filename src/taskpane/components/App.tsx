import * as React from "react";
import Progress from "./Progress";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
/* global Button, console, Header, HeroList, HeroListItem, Office, Progress */

//= ui-Dropdown 
import { Dropdown, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown'
const dropdownStyles: Partial<IDropdownStyles>  = {
  dropdown: { width: 80 },
};
const durationOptions: IDropdownOption[] = [
  { key: 'duration60', text: '60' },
  { key: 'duration40', text: '40' },
  { key: 'duration30', text: '30' },
  { key: 'duration25', text: '25' },
];
const intervalOptions: IDropdownOption[] = [
  { key: 'interval10', text: '10' },
  { key: 'interval5', text: '5' },
  { key: 'interval1', text: '1' },
];
function _dropdownSelected(e, selectedOption): void {
  this.setState({
    duration: selectedOption,
    interval: selectedOption
  });
}

//= ui-Button
import { PrimaryButton } from 'office-ui-fabric-react';
function _alertClicked(): void {
  alert('Seted!');
}



export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  duration: 0,
  interval: 0,
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      duration: 0,
      interval: 0,
    };
  }

  componentDidMount() {
    this.setState({
    });
  }

  click = async () => {
    /**
     * Insert your PowerPoint code here
     */
    Office.context.document.setSelectedDataAsync(
      "Hello World!",
      {
        coercionType: Office.CoercionType.Text
      },
      result => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/yasashi_DX.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Dropdown 
          label="Duration"
          options={durationOptions}
          styles={dropdownStyles}
        />
        <Dropdown 
          label="Interval"
          options={intervalOptions}
          styles={dropdownStyles}
        />
        <PrimaryButton
          text="Set"
          onClick={this.click}
          allowDisabledFocus
        />
      </div>
    );
  }
}
