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
  };

public async componentDidMount() {
    await this.fetchData();
}

/**
 * fetchData
 */
public async fetchData() {
  let web = Web(this.props.webURL);
  const item: any[] = await web.lists.getByTitle("EmployeeDetails").items.select("*", "Employee_x0020_Name/Title").expand("Employee_x0020_Name/ID").get();
  console.log(itmes);
  this.setState({ Items: itmes});
  let html = await this.getHTML(items);
  this.setState({HTML: html});
}
public findData = (id): void => {
  //this.fetch();
  var itemID = id;
  var allitems = this.state.Items;
  var allitemsLength = allitems.length;
  if (allitemsLength > 0) {
    for (let i = 0; i < allitemsLength; i++) {
      if (itemID == allitems[i].Id) {
        this.setState({
          ID:itemID,
          EmployeeName: allitems[i].Employee_x0020_Name.Title,
          EmployeeNameId:allitems[i].Employee_x0020_NameId,
          HireDate: new Date(allitems[i].HireDate),
          JobDescription:allitems[i].Job_x0020_Description
        });
      }

    }
  }
}

/**
 * getHTML
 */
public async getHTML() {
 var tabledata= <table className={styles.table}>
  <thead>
    <tr>
      <th>Employee Name</th>
       <th>Hire Date</th>
      <th>Job Description</th>
   </tr>
  </thead>
  <tbody>
    {item && Items.map((item, i) => {
      return [
        <tr key={i} onClick={()=>this.findData(item.ID)}>
        <td>{item.Employee_x0020_Name.Title}
        </td>
        <td>
          {FormatDate(item.HireDate)}
        </td>
        <td>
          {item.Job_x0020_Description}
        </td>
        </tr>
      ]
    })}
  </tbody>
 </table>;
 return await tabledata;
}

/**
 * _getPeaplePickerItems = async
items: any[] */
public _getPeaplePickerItems = async (items: any[]) {
  if (items.length > 0) {
    this.setState({ EmployeeName: items[0].text});
    this.setState({EmployeeNameId: items[0].id});
  }
  else{
    //ID=0;
    this.setState({ EmplyeeNameId: ""});
    this.setState({ EmplyeeName: ""});
  }
}

/**
 * onchange
value, stateValue */
public onchange(value, stateValue) {
  let state = {};
  state[stateValue] = value;
  this.setState(state);
}

private async SaveDate() {
  let web = Web(this.props.webURL);
  await web.lists.getByTitle("EmployeeDetails").items.add({

    Employee_x0020_NameId: this.state.EmployeeNameId,
    HireDate: new Date(this.state.HireDate),
    Job_x0020_Description: this.state.JobDescripton,
  }).then(i => {
    console.log(i);
  });
  alert("Created Successfully");
  this.setState({EmployeeName:"",HireDate:null, JobDescription:""});
  this.fetchData();
}
private async UpdateData() {
  let web = Web(this.props.webURL);
  await web.lists.getByTitle("EmployeeDetails").items.getById(this.state.ID).update({

    Employee_x0020_NameId: this.state.EmployeeNameId,
    HireDate: new Date(this.state.HireDate),
    Job_x0020_Description: this.state.JobDescripton,
  }).then(i => {
    console.log(i);
  });
  alert("Updated Successfully");
  this.setState({EmployeeName:"",HireDate:null, JobDescription:""});
  this.fetchData();
}

private async DeleteData() {
  let web = Web(this.props.webURL);
  await web.lists.getByTitle("EmployeeDetails").items.getById(this.state.ID).delete()
  .then(i => {
    console.log(i);
  });
  alert("Deleted Successfully");
  this.setState({EmployeeName:"",HireDate:null, JobDescription:""});
  this.fetchData();
}

  public render(): React.ReactElement<IKcrudProps> {
    return(
      <div>
      <h1>CRUD Operations in SharePoint using SPFx with ReactJs</h1>
      {this.state.HTML}
        <div className={styles.button}>
        <div><PrimaryButton text="Create" onClick={()=> this.SaveData()}/></div>
        <div><PrimaryButton text="Update" onClick={()=> this.UpdateData()}/></div>
        <div><PrimaryButton text="Delete" onClick={()=> this.DeleteData()}/></div>
        </div>
        <div>
        <form>
        <div>
        <Label>EMployee Name</Label>
        <PeoplePicker
        context={this.props.context}
        personSelectionLimit={1}
        // defaultSelectedUsers={this.state.EmployeeName===""?[]:this.state.EmployeeName}
        required={false}
        onChange={this._getPeoplePickerItems}
        defaultSelectedUsers={[this.state.EmployeeName?this.state.EmployeeName:""]}
        ShowHiddenInUI={false}
        principalTypes={[PrincipalType.User]}
        resolveDelay={1000}
        ensureUser={true}
        />
        </div>
        <div>
        <Label>HireDate</Label>
        <DatePicker maxDate={new Date()} allowTextInput={false} strings = {DatePickerStrings} value={this.state.HireDate}
        />
        </div>
        <div>
          <Label>Job Description</Label>
          <TextField value={this.state.JobDescripton} multiline onChanged={(value)=> this.onchange(value, "jobDescription")}
          />
        </div>
        </form>
          <div className = { styles.kcrud }>
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
      </div>
    );
  }
}
