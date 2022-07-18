// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { connect } from 'react-redux';
import { withTranslation, WithTranslation } from "react-i18next";
import { Menu, MoreIcon } from '@fluentui/react-northstar';
import { getBaseUrl } from '../../configVariables';
import * as microsoftTeams from "@microsoft/teams-js";
import { duplicateDraftNotification, cancelSentNotification, emailUnreadsNotification, getAppSettings } from '../../apis/messageListApi';
import { selectMessage, getMessagesList, getDraftMessagesList } from '../../actions';
import { TFunction } from "i18next";

export interface OverflowProps extends WithTranslation {
    message?: any;
    styles?: object;
    title?: string;
    selectMessage?: any;
    getMessagesList?: any;
    getDraftMessagesList?: any;
}

export interface OverflowState {
    menuOpen: boolean;
    enableEmailFallback: boolean;
}

export interface ITaskInfo {
    title?: string;
    height?: number;
    width?: number;
    url?: string;
    card?: string;
    fallbackUrl?: string;
    completionBotId?: string;
}

class Overflow extends React.Component<OverflowProps, OverflowState> {
    readonly localize: TFunction;
    enableEmailFallback: boolean;
    constructor(props: OverflowProps) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            menuOpen: false,
            enableEmailFallback: false,
        };
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        this.getAppSettings().then(() => {
            
            
        });
    }

    public render(): JSX.Element {
        let shouldNotShowCancel;
        if (this.props.message != undefined && this.props.message.status != undefined) {
            const status = this.props.message.status.toUpperCase();
            shouldNotShowCancel = status === "SENT" || status === "UNKNOWN" || status === "FAILED" || status === "CANCELED" || status === "CANCELING";
        }

        const items = [
            {
                key: 'more',
                icon: <MoreIcon outline={true} />,
                menuOpen: this.state.menuOpen,
                active: this.state.menuOpen,
                indicator: false,
                menu: {
                    items: [
                        {
                            key: 'status',
                            content: this.localize("ViewStatus"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                let url = getBaseUrl() + "/viewstatus/" + this.props.message.id + "?locale={locale}";
                                this.onOpenTaskModule(null, url, this.localize("ViewStatus"));
                            }
                        },
                        {
                            key: 'duplicate',
                            content: this.localize("Duplicate"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                this.duplicateDraftMessage(this.props.message.id).then(() => {
                                    this.props.getDraftMessagesList();
                                });
                            }
                        },
                        {
                            key: 'cancel',
                            content: this.localize("Cancel"),
                            hidden: shouldNotShowCancel,
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                this.cancelSentMessage(this.props.message.id).then(() => {
                                    this.props.getMessagesList();
                                });
                            }
                        },
                        {
                            //new option to notify unreads using email
                            key: 'emailunreads',
                            content: this.localize("emailunreads"),
                            hidden: !(this.state.enableEmailFallback),
                            
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                this.emailunreads(this.props.message.id).then(() => {
                                    //TODO: Update the sent list with a message saying email was triggered to notify unread messages
                                    this.props.getMessagesList();
                                });
                            }
                        },
                    ],
                },
                onMenuOpenChange: (e: any, { menuOpen }: any) => {
                    this.setState({
                        menuOpen: menuOpen
                    });
                },
            },
        ];

        return <Menu className="menuContainer" iconOnly items={items} styles={this.props.styles} title={this.props.title} />;
    }

    private onOpenTaskModule = (event: any, url: string, title: string) => {
        let taskInfo: ITaskInfo = {
            url: url,
            title: title,
            height: 530,
            width: 1000,
            fallbackUrl: url,
        };
        let submitHandler = (err: any, result: any) => {
        };
        microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    }

    private duplicateDraftMessage = async (id: number) => {
        try {
            await duplicateDraftNotification(id);
        } catch (error) {
            return error;
        }
    }

    // get the app configuration values and set targeting mode from app settings
    private getAppSettings = async () => {
        let response = await getAppSettings();
            if (response.data) {
                this.setState({
                    enableEmailFallback: response.data.enableEmailFallback, 
            });
        }
    }

    private cancelSentMessage = async (id: number) => {
        try {
            await cancelSentNotification(id);
        } catch (error) {
            return error;
        }
    }

    //calls the emailUnreadsNotification to send emails to all users who not read the message with id
    private emailunreads = async (id: number) => {
        try {
            await emailUnreadsNotification(id);
        } catch (error) {
            return error;
        }
    }
}

const mapStateToProps = (state: any) => {
    return { messagesList: state.messagesList };
}

const overflowWithTranslation = withTranslation()(Overflow);
export default connect(mapStateToProps, { selectMessage, getMessagesList, getDraftMessagesList })(overflowWithTranslation);
