import * as React from 'react';
import styles from './NewForm.module.scss';
import { INewFormProps } from './INewFormProps';
import { INewFormStates } from './INewFormStates';
import { escape } from '@microsoft/sp-lodash-subset';
import { Container, Typography, FormControl, TextField, Select, MenuItem, Button } from '@material-ui/core';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


export default class NewForm extends React.Component<INewFormProps, INewFormStates, any> {

  constructor(props) {
    super(props);

    this.state = {
      task: '',
      description: '',
      priority: '',
      dueDate: null,
      taskExecutor: [],
      emailTaskExecutor: ''
    };
  }

  componentDidMount() {
    this.fetchData();
  }

  async fetchData() {
    let web = Web(this.props.webURL);
    const items: any[] = await web.lists.getById('8414a250-0699-4efa-afcc-f4a34b89498c').items.get();
    console.log('items', items);
  }

  public handleChange = (e: {target: {name: any; value: any; }; }) => {
    console.log(e.target.value);
    const newState = { [e.target.name]: e.target.value } as Pick<INewFormStates, keyof INewFormStates>;
    this.setState(newState);
  }

  // public componentDidUpdate(prevProps: Readonly<INewFormProps>, prevState: Readonly<INewFormStates>, snapshot?: any): void {
  //   console.log('Updated', this.state);
  // }

  public handleSelectChange = (e) => {
    console.log('e.target',e.target.value);
    this.setState({ priority: e.target.value });
  }

  public getPeoplePickerItems = (items: any[]) => {
    console.log('Items:', items);
    this.setState({taskExecutor: items, emailTaskExecutor: items[0].secondaryText});
  }

  public ResetForm  = () => {
    this.setState({ task: '', description: '', priority: '', dueDate: null, taskExecutor: [], emailTaskExecutor: ''}); 
  }

  public AddItem = () => {
    console.log('state', this.state);
    let web = Web(this.props.webURL);
    web.lists.getById('8414a250-0699-4efa-afcc-f4a34b89498c').items.add({
      Title: this.state.task,
      Description: this.state.description,
      Priority: this.state.priority,
      Due_x0020_date: this.state.dueDate,
      Task_x0020_ExecutorId: this.state?.taskExecutor[0]?.id,
      Email_x0020_Task_x0020_Executor: this.state.emailTaskExecutor
    }).then(AddResult => {
      console.log('Create AddResult', AddResult);
      let taskId = AddResult.data.ID;
      web.lists.getById('8414a250-0699-4efa-afcc-f4a34b89498c').items.getById(taskId).update({
        NewID: {
          "__metadata": { "type": "SP.FieldUrlValue" },
          "Description": taskId,
          "Url": "https://projects1.sharepoint.com/sites/Development/Alona/SitePages/EditForm.aspx?FormID=" + taskId
        }
      }).then(UpdateResult  => {
        if(onclick) {
          this.ResetForm();
        }
      })
    });
    alert("Created Successfully");
    console.log('saved-state', this.state);
  }

  public render(): React.ReactElement<INewFormProps> {
    return (
      <Container maxWidth="sm">
        <Typography variant="h6" style={{ textAlign: 'center', marginTop: '10px', marginBottom: '10px' }}>New Task</Typography>
        <FormControl style={{ marginTop: '20px' }}>
          <TextField label="Task" name="task" value={this.state.task} onChange={this.handleChange} variant="outlined" style={{ marginTop: '13px', width: '500px' }}/>
          <TextField label="Task Description" name="description" value={this.state.description} onChange={this.handleChange} variant="outlined" multiline rows={3} style={{ marginTop: '13px', width: '500px' }}/>
          <Select
            label="Priority"
            name="priority"
            value={this.state.priority ? this.state.priority : 'Low'}
            onChange={(e) => {this.handleSelectChange(e)}}
            variant="outlined" 
            style={{ marginTop: '13px', width: '500px' }}
          >
            <MenuItem value="High">High</MenuItem>
            <MenuItem value="Medium">Medium</MenuItem>
            <MenuItem value="Low">Low</MenuItem>
          </Select>
          <TextField
            id="date"
            variant="outlined"
            label="Due Date"
            type="date"
            name="dueDate"
            // value={this.state.dueDate ? this.state.dueDate : (new Date().toJSON().slice(0,10))}
            value={this.state.dueDate ? this.state.dueDate : ''}
            InputLabelProps={{
              shrink: true,
            }}
            style={{ marginTop: '13px', width: '500px' }}
            onChange={this.handleChange}
          />
          <PeoplePicker
            context={this.props.context as any}
            titleText="Task Executor"
            groupName={''}
            personSelectionLimit={1}
            required={false}
            showHiddenInUI={false}
            defaultSelectedUsers={this.state.taskExecutor}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            ensureUser={true}
            onChange={this.getPeoplePickerItems}
          />
        </FormControl>
        <div style={{ marginTop: '20px' }}>
          <Button style={{ width: '83px', marginRight: '5px'}} variant="outlined" color="primary" onClick={this.AddItem}>Save</Button>
          <Button variant="outlined" color="secondary" onClick={this.ResetForm}>Cancel</Button>
        </div>
      </Container>

    );
  }
}
