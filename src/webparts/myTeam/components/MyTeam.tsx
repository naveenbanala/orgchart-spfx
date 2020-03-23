import * as React from 'react';
import styles from './MyTeam.module.scss';
import { IMyTeamProps } from './IMyTeamProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Services from '../../../utilities/Services'
import * as constants from '../../../utilities/Constants'
import { ISPState, ISPUser, ISPUsers } from '../../../Model/DataModel'
import LivePersonaCard from '../../LivePersonaCard'


export default class MyTeam extends React.Component<IMyTeamProps, { MyDetails, UserPeerCollection, UserManager, UserDirectReportCollection }> {
  constructor(IMyTeamProps) {
    super(IMyTeamProps);
    this.state = {
      UserPeerCollection: [],
      UserManager: [],
      UserDirectReportCollection: [],
      MyDetails: null
    }
  }

  public componentDidMount() {
    let users = []
    Services.getUserInfo(this.props.context, constants.MyDetails).then(resp => {
      if (!resp.error) {
        this.setState({
          MyDetails: resp
        })
      }
    })
    Services.getUserInfo(this.props.context, constants.MyManager).then(resp => {
      if (!resp.error) {
        this.setState({
          UserManager: [...this.state.UserManager, resp]
        }, () => {
          Services.getUserInfo(this.props.context, constants.UserDirectReports.replace('{id}', resp.id)).then(resp => {
            if (!resp.error) {
              resp.value && resp.value.length > 0 && resp.value.map(eachUser => {
                if (eachUser.displayName != this.state.MyDetails.displayName)
                  this.setState({
                    UserPeerCollection: [...this.state.UserPeerCollection, eachUser]
                  })
              })
            }
          })
        })
      }
    })

    Services.getUserInfo(this.props.context, constants.MyDirectReports).then(resp => {
      if (resp.value && resp.value.length > 0) {
        resp.value.map(eachuser => {
          this.setState({
            UserDirectReportCollection: [...this.state.UserDirectReportCollection, eachuser]
          })
        })
      }
    })

  }

  private FilterUsersBasedonProps(manager, directReports, peers) {
    if (manager && peers && directReports) {
      return [...this.state.UserManager, ...this.state.UserDirectReportCollection, ...this.state.UserPeerCollection]
    } else if (manager && peers && !directReports) {
      return [...this.state.UserManager, ...this.state.UserPeerCollection]
    } else if (manager && !peers && directReports) {
      return [...this.state.UserManager, ...this.state.UserDirectReportCollection]
    } else if (!manager && peers && directReports) {
      return [...this.state.UserDirectReportCollection, ...this.state.UserPeerCollection]
    } else if (manager && !peers && !directReports) {
      return [...this.state.UserManager]
    } else if (!manager && peers && !directReports) {
      return [...this.state.UserPeerCollection]
    } else if (!manager && !peers && directReports) {
      return [...this.state.UserDirectReportCollection]
    } else {
      return []
    }
  }

  public render(): React.ReactElement<IMyTeamProps> {

    let AllUsers = this.FilterUsersBasedonProps(this.props.checkboxManagers, this.props.checkboxDirectReports, this.props.checkboxPeers)
    console.log("[user collection]", AllUsers)
    return (
      <div className={""}>
        <div>
          <h3>{this.props.description}</h3>
        </div>
        {AllUsers && AllUsers.length > 0 && AllUsers.map(user => <LivePersonaCard {...user} context={this.props.context} />)}
      </div>
    );
  }
}
