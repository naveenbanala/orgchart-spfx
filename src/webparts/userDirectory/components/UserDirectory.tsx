import * as React from 'react';
import styles from './UserDirectory.module.scss';
import { IUserDirectoryProps } from './IUserDirectoryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISPState, ISPUser, ISPUsers } from '../../../Model/DataModel'
import Services from '../../../utilities/Services'
import LivePersonaCard from '../../LivePersonaCard'

export default class UserDirectory extends React.Component<IUserDirectoryProps, ISPState> {

  constructor(IUserDirectoryProps) {
    super(IUserDirectoryProps);
    this.state = {
      User: null,
      UserCollection: null
    }
  }

  private GetAllOrgUsers = () => {
    Services.getAllUsers(this.props.context).then(resp => {
      this.setState({
        UserCollection: resp
      })
    })
  }

  public componentDidMount() {
    this.GetAllOrgUsers()
  }

  public render(): React.ReactElement<IUserDirectoryProps> {
    let alluser = { ...this.state.UserCollection }
    console.log("[user collection]", alluser)
    return (
      <div className={""}>
        {alluser.value && alluser.value.length > 0 && alluser.value.map(user => <LivePersonaCard {...user} context={this.props.context} />)}
      </div>
    );
  }
}
