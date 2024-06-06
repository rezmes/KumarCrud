import * as React from 'react';
import styles from './Kcrud.module.scss';
import { IKcrudProps } from './IKcrudProps';

import {DatePicker, IDatePickerStrings} from 'office-ui-fabric-react/lib/DatePicker';
import {PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Label } from "office-ui-fabric-react/lib/Label";
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { escape } from '@microsoft/sp-lodash-subset';


export interface IStates {
Items: any;
ID: any;
HireDate: any;
EmployeeName: any;
EmployeeId: any;
JobDescripton: any;
HTML: any;
}


export default class Kcrud extends React.Component < IKcrudProps, IStates > {
  constructor(props) {
    super(props);
    this.state = {
      Items: [],
      ID: 0,
      HireDate: null,
      EmployeeName: "",
      EmployeeId: 0,
      JobDescripton: "",
      HTML: []
  }
  public render(): React.ReactElement<IKcrudProps> {
    return(
      <div>
      <h1>CRUD Operations in SharePoint using SPFx with ReactJs</h1>
      {this.state.HTML}
      <div className={styles.button}>
      <div className = { styles.kcrud } >
  <div className={styles.container}>
    <div className={styles.row}>
      <div className={styles.column}>
        <span className={styles.title}>Welcome to SharePoint!</span>
        <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
        <p className={styles.description}>{escape(this.props.description)}</p>
        <a href='https://aka.ms/spfx' className={styles.button}>
          <span className={styles.label}>Learn more</span>
        </a>
      </div>
    </div>
  </div>
      </div >
      </div>
    );
  }
}
