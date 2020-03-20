import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISPState, ISPUser, ISPUsers } from '../Model/DataModel'
import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { Stack, Label, Link, replaceElement, List } from 'office-ui-fabric-react/lib';
import { HoverCard, IExpandingCardProps, IHoverCardProps } from 'office-ui-fabric-react/lib/HoverCard';
import Service from '../utilities/Services'
import * as constants from '../utilities/Constants'
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Separator } from 'office-ui-fabric-react/lib/Separator';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export default class LivePersonaCard extends React.Component<ISPUser, { User, MgrUser, EmpDR }>{
    constructor(props: ISPUser) {
        super(props);
        this.state = {
            User: null,
            MgrUser: null,
            EmpDR: null
        }

    }
    componentDidMount() {
        Service.getUserphoto(this.props.context, this.props.id).then(res => {
            if (res.status == 200) {
                console.log("[image resp]", res)
                res.arrayBuffer().then((buffer) => {
                    var base64Flag = 'data:image/jpeg;base64,';
                    var imageStr = this.arrayBufferToBase64(buffer);
                    this.setState({
                        EmpDR: base64Flag + imageStr
                    })
                })
            } else {
                this.setState({
                    EmpDR: ""
                })
            }
        })
    }

    public render(): JSX.Element {

        let examplePersona: IPersonaSharedProps = {
            imageUrl: this.state.EmpDR,
            imageInitials: this.getInitials(this.props.displayName),
            text: this.props.displayName,
            secondaryText: this.props.jobTitle
        };

        const expandingCardProps: IExpandingCardProps = {
            onRenderCompactCard: this._onRenderCompactCard,
            onRenderExpandedCard: this._onRenderExpandedCard,
            renderData: { "persona": this.props, "photo": this.state.EmpDR }
        };
        return (
            <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg3" >
                <HoverCard expandingCardProps={expandingCardProps} onCardHide={this._onCardHide}>
                    <Persona
                        {...examplePersona}
                    />
                </HoverCard>
            </div>
        )
    }

    private _onCardHide = () => {
        this.setEmpManagerNull()
        console.log("[_onCardHide]");
    }

    private setEmpManagerNull = () => {
        this.setState({
            MgrUser: null
        })
    }


    private _onRenderCompactCard = (item: any): JSX.Element => {
        let managerPersona: IPersonaSharedProps = {
            imageUrl: this.state.EmpDR,
            imageInitials: this.getInitials(item.persona.displayName),
            text: item.persona.displayName,
            secondaryText: item.persona.jobTitle,
            tertiaryText: item.persona.department,
        };
        return (
            <div className="ms-Grid">
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                        <Persona
                            {...managerPersona}
                            size={PersonaSize.size72}
                        />
                    </div>
                </div>
            </div>
        );
    }

    private _onRenderExpandedCard = (item: any): JSX.Element => {
        this.getManagerDetails(item.persona.id)
        // this.getDirectReportsDetails(item.persona.id)

        let managerPersona: IPersonaSharedProps = this.state.MgrUser ? {
            imageUrl: this.state.MgrUser.photo,
            imageInitials: this.getInitials(this.state.MgrUser.displayName),
            text: this.state.MgrUser.displayName,
            secondaryText: this.state.MgrUser.jobTitle,
            tertiaryText: this.state.MgrUser.department,
        } : null;

        // let directreportsPersona: IPersonaSharedProps = this.state.EmpDR ? {
        //     imageUrl: this.state.EmpDR.photo,
        //     imageInitials: this.getInitials(this.state.EmpDR.displayName),
        //     text: this.state.EmpDR.displayName,
        //     secondaryText: this.state.EmpDR.jobTitle,
        //     tertiaryText: this.state.EmpDR.department,
        // } : null;

        return (
            <div className={""}>
                <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <h3><Label>Contact<Icon iconName="ChevronRight" className="ms-Icon" /></Label></h3>
                    </div>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2"><Icon iconName="Mail" className="ms-Icon" /></div>
                        {item.persona.mail ? <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"><Label>{item.persona.mail}</Label></div> : "--"}
                    </div>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2"><Icon iconName="Phone" className="ms-Icon" /></div>
                        {item.persona.mobilePhone ? <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"><Label>{item.persona.mobilePhone}</Label></div> : "--"}

                    </div>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2"><Icon iconName="MapPin" className="ms-Icon" /></div>
                        {item.persona.officeLocation ? <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"><Label>{item.persona.officeLocation}</Label></div> : "--"}
                    </div>
                </div>
                <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <Separator alignContent="start">
                            <h3><Label>Manager<Icon iconName="ChevronRight" className="ms-Icon" /></Label></h3>
                            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                                {this.state.MgrUser ? <Persona
                                    {...managerPersona}
                                /> : <h4>none</h4>}
                            </div>
                        </Separator>
                    </div>
                </div>
                {/* <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <Separator alignContent="start">
                            <h3><Label>Reports to<Icon iconName="ChevronRight" className="ms-Icon" /></Label></h3>
                            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                                {this.state.EmpDR ? <Persona
                                    {...directreportsPersona}
                                /> : <h4>none</h4>}
                            </div>
                        </Separator>
                    </div>
                </div> */}
            </div>
        );

    }

    private getManagerDetails = (id: any): any => {
        if (!this.state.MgrUser) {
            Service.getUserInfo(this.props.context, constants.UserManager.replace("{id}", id)).then(mgResp => {
                console.log(mgResp)
                if (!mgResp.error && !mgResp.value) {
                    Service.getUserphoto(this.props.context, mgResp.id).then(res => {
                        if (res.status == 200) {
                            console.log("[image resp]", res)
                            res.arrayBuffer().then((buffer) => {
                                var base64Flag = 'data:image/jpeg;base64,';
                                var imageStr = this.arrayBufferToBase64(buffer);
                                mgResp.photo = base64Flag + imageStr
                            })
                        } else {
                            mgResp.photo = ""
                        }
                        console.log("[manager details]", mgResp)
                        this.setState({
                            MgrUser: mgResp
                        })
                    })

                }
            })
        }
    }
    private getDirectReportsDetails = (id: any): any => {
        if (!this.state.EmpDR) {
            Service.getUserInfo(this.props.context, constants.UserDirectReports.replace("{id}", id)).then(mgResp => {
                console.log(mgResp)
                if (!mgResp.error && !mgResp.value) {

                    Service.getUserphoto(this.props.context, mgResp.id).then(res => {
                        if (res.status == 200) {
                            console.log("[image resp]", res)
                            res.arrayBuffer().then((buffer) => {
                                var base64Flag = 'data:image/jpeg;base64,';
                                var imageStr = this.arrayBufferToBase64(buffer);
                                mgResp.photo = base64Flag + imageStr
                            })
                        } else {
                            mgResp.photo = ""
                        }
                        console.log("[manager details]", mgResp)
                        this.setState({
                            EmpDR: mgResp
                        })
                    })

                }
            })
        }
    }

    private getInitials = (name: any) => {
        let initials = name.match(/\b\w/g) || [];
        initials = ((initials.shift() || '') + (initials.pop() || '')).toUpperCase();
        return initials
    }

    private compare(a: any, b: any) {
        if (a.displayName < b.displayName) {
            return -1;
        }
        if (a.displayName > b.displayName) {
            return 1;
        }
        return 0;
    }
    private arrayBufferToBase64(buffer) {
        var binary = '';
        var bytes = [].slice.call(new Uint8Array(buffer));
        bytes.forEach((b) => binary += String.fromCharCode(b));
        return window.btoa(binary);
    };
}
