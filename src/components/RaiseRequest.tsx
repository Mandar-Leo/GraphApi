import React, { ChangeEvent, FormEvent, FormEventHandler, useContext, useEffect, useRef, useState } from "react";
import { Dialog, Text, Button, Input, Label, Dropdown, Option, Checkbox, Image, makeStyles, FluentProvider, makeResetStyles, tokens, Field, Textarea, InputOnChangeData, Theme, OptionOnSelectData, SelectionEvents, DialogTrigger, DialogSurface, DialogBody, DialogTitle, DialogContent, DialogActions } from "@fluentui/react-components";
import { app, Context } from "@microsoft/teams-js";
import i18n from "../i18n";
import { SearchBox } from "@fluentui/react-search-preview";
// @ts-ignore
import { SearchBoxChangeEvent } from "../../src/SearchBox";
// @ts-ignore
import detectBrowserLanguage from "detect-browser-language";
import { TeamsUserCredentialAuthConfig, TeamsFx, IdentityType } from "@microsoft/teamsfx";

import { I18nextProvider } from "react-i18next";
import { API_SCOPES, LIST_ID, SITE_ID } from "../Constants";
import { TeamsFxProvider } from '@microsoft/mgt-teamsfx-provider';
import { TeamsUserCredential, createMicrosoftGraphClient } from "@microsoft/teamsfx";

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider, TokenCredentialAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { TeamsFxContext } from "./Context";
import { userInfo } from "os";
import { LoginType, ProviderState, Providers, registerComponent } from "@microsoft/mgt-element";
import { Msal2Provider, Msal2Config, Msal2PublicClientApplicationConfig } from "@microsoft/mgt-msal2-provider";
import { PeoplePicker, Person } from "@microsoft/mgt-react";
import User from "./User";
import { createRequest } from "../services/GraphAPIs";
import { Route, Router, useNavigate, useNavigation } from "react-router-dom";
// import TeamsDetails from "./TeamsDetails";


interface IRaiseRequestProps {
    // accessToken: string;
    // themeName: Theme
    graphClient: Client
}

interface IOneCollabUser {
    displayName: string
    givenName: string
    jobTitle: string
    mail: string
}

const useStackClassName = makeResetStyles({
    display: "flex",
    paddingLeft: "1.5rem",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalL,
});

Providers.globalProvider = new Msal2Provider({
    clientId: process.env.REACT_APP_CLIENT_ID as string,
    scopes: API_SCOPES
});

export function RaiseRequest(props: IRaiseRequestProps) {

    const { teamsUserCredential } = useContext(TeamsFxContext);
    const [valid, setValid] = useState(true);

    const [title, setTitle] = useState("");
    const [description, setDescription] = useState("");
    const [requester, setRequester] = useState("");
    const [owner, setOwner] = useState("");
    const [coOwners, setCoOwners] = useState("");
    const [justification, setJustification] = useState("");
    const [lifetime, setLifetime] = useState("");
    const [privacy, setPrivacy] = useState("");
    const [sensitivity, setSensitivity] = useState("");

    const [searchedUsers, setSearchedUsers] = useState([] as IOneCollabUser[])

    const [openValidationDialog, setOpenValidationDialog] = useState(false);
    const [validationDialogTitle, setValidationDialogTitle] = useState("");
    const [validationDialogMessage, setValidationDialogMessage] = useState("");
    const [openTeamsDialog, setOpenTeamsDialog] = useState(false);
    const [teamsDialogTitle, setTeamsDialogTitle] = useState("");
    const [teamsDialogMessage, setTeamsDialogMessage] = useState("");
    const [teamDetails, setTeamDetails] = useState("");
    const [nameVerified, setNameVerified] = useState(false);

    const refRequester: any = useRef();
    const refOwner: any = useRef();
    const refCoOwners: any = useRef();

    // let foo = Route()
    // console.log("navigation.state", props)

    useEffect(() => {



        let requests = props.graphClient.api(`/sites/${SITE_ID}/lists/${LIST_ID}/items`)
            .expand("fields")
            // .filter(`fields/Requester eq '${userPrincipalName}'`)
            .get()
            .then((value: any) => {
                console.log("requests", value)
            });

        console.log('In getRequests - requests...');
        console.log(requests);

    }, []);

    function validateEntry(entry: string): boolean {
        console.log("validateEntry", entry)
        if (entry == null || entry === undefined || entry.trim() === "") {
            return false;
        } else {
            return true;
        }
    }

    function makeListEntry(graphClient: any, title: string, description: string, requester: string, owner: string, coOwner: string, justification: string, lifetime: string, autoDelete: boolean, privacy: string) {

        console.log("in makeListEntry");
        console.log(requester, owner, coOwner);

        createRequest(graphClient, title, description, requester, owner, coOwner, justification, lifetime, true, privacy)
            .then((result) => {

                console.log("in createRequest then");
                console.log(result);

                setTitle("Test Title - ");
                setDescription("Test long description for 50 longer characters almost");
                setRequester("");
                setOwner("");
                setCoOwners("");
                setJustification("Test long justification for 50 longer characters almost");
                setLifetime("");
                setPrivacy("")
                setSensitivity("");
                setOpenValidationDialog(false);
                setValidationDialogTitle("Status Update");
                setValidationDialogMessage(`<p>Your request is successfully created. Please save the ID '${result.id}' for your records.</p>
                You can safely close this dialog and navigate away from this page now!`);
            });
    }

    const handleSubmitRequest = async () => {
        console.log("in handleSubmitRequest");

        let selectedRequester: string[] = await refRequester.current.getSelectedUsers();
        let selectedOwner: string[] = await refOwner.current.getSelectedUsers();
        let selectedCoOwners: string[] = await refCoOwners.current.getSelectedUsers();

        let errorMessage: string = "";
        let nameRegEx: RegExp = new RegExp(`[~"#%&*:<>?\/\\{|}.]`);

        errorMessage +=
            (title.search(nameRegEx) > 0
                || validateEntry(title) === false
                || title.length < 10
                || title.length > 64)
                ? i18n.t('raiseRequest.validationMessage.titleName')
                : "";

        // errorMessage +=
        //   (!this.state.teamNameVerified)
        //     ? i18n.t('raiseRequest.validationMessage.teamNameVerified')
        //     : "";

        errorMessage +=
            (validateEntry(description) === false
                || description.length < 50
                || description.length > 100)
                ? i18n.t('raiseRequest.validationMessage.description')
                : "";

        errorMessage +=
            (selectedRequester.length == 0)
                ? i18n.t('raiseRequest.validationMessage.requester')
                : "";

        errorMessage +=
            (selectedOwner.length == 0)
                ? i18n.t('raiseRequest.validationMessage.owner')
                : "";

        errorMessage +=
            (selectedCoOwners.length == 0)
                ? i18n.t('raiseRequest.validationMessage.coOwner')
                : "";

        errorMessage +=
            (validateEntry(justification) === false
                || justification.length < 50
                || justification.length > 100)
                ? i18n.t('raiseRequest.validationMessage.businessJustification')
                : "";

        errorMessage +=
            (validateEntry(lifetime) === false)
                ? i18n.t('raiseRequest.validationMessage.lifeTime')
                : "";

        errorMessage +=
            (validateEntry(privacy) === false)
                ? i18n.t('raiseRequest.validationMessage.privacyLevel')
                : "";

        errorMessage +=
            (validateEntry(sensitivity) === false)
                ? i18n.t('raiseRequest.validationMessage.sensitivityLevel')
                : "";

        if (errorMessage !== "") {
            // alert("invalid entries");
            setOpenValidationDialog(false);
            setValidationDialogTitle("Let's address the following items:");
            setValidationDialogMessage(`<ul>${errorMessage}</ul>`);
        } else {
            console.log("in handleSubmitRequest");
            console.log("values: " + requester + " --- " + owner + " --- " + coOwners);

            //     //   this.setState(
            //     //     {
            //     //       requester: requester,
            //     //       owner: owner,
            //     //       coOwner: coOwner,
            //     //     },
            //     //     this.makeListEntry
            //   );
            makeListEntry(props.graphClient, title, description, selectedRequester[0], selectedOwner[0], selectedCoOwners.join(";"), justification, lifetime, false, privacy);
        }
        setOpenValidationDialog(true);
    }

    async function searchUser(strSearchText: string) {
        const url = `https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${strSearchText}')`;

        return props.graphClient
            .api(url)
            .get();
    }

    // function handleSubmitRequest() {
    //     console.log("title:", title, "description:", description, "requester:", requester, "owner:", owner, "coOwner:", coOwner, "justification:", justification);
    // }

    function onChange(
        ev: SearchBoxChangeEvent,
        data: InputOnChangeData
    ) {
        console.log(data);
        // if (data.value.length <= 20) {
        //     setRequesterValue(data.value);
        //     setValid(true);
        // } else {
        //     setValid(false);
        // }
    };

    // const useStyles = makeStyles({
    //     TitleStyle: {
    //         fontWeight: "bold",
    //         color: this.setStyles('white', 'white', 'black'),
    //         backgroundColor: this.setStyles('black', 'black', '#F5F5F5')
    //     },
    //     FlexStyle: {
    //         paddingLeft: "1.5rem",
    //         color: this.setStyles('white', 'white', 'black'),
    //         backgroundColor: this.setStyles('black', 'black', '#F5F5F5'),//#292929
    //     },
    //     DialogStyle: {
    //         color: this.setStyles('white', 'white', 'black'),
    //         backgroundColor: this.setStyles('black', 'black', '#F5F5F5'),
    //         height: '75%',
    //         width: '75%'
    //     },
    //     FormFieldStyle: {
    //         width: "65%",
    //     },
    //     InputStyle: {
    //         color: this.setStyles('white', 'white', 'black'),
    //         backgroundColor: this.setStyles('black', 'black', '#F5F5F5')
    //         // border: '1px',
    //         // borderColor: this.setStyles('black', 'white', 'black')
    //     },
    //     DropDownStyle: {
    //         // color: this.setStyles('black', 'white', 'black'),
    //         // backgroundColor: this.setStyles('white', 'black', '#F5F5F5'),
    //     },
    //     LabelStyle: {
    //         marginBottom: ".25rem",
    //         fontWeight: "bold",
    //         color: this.setStyles('white', 'white', 'black'),
    //         backgroundColor: this.setStyles('black', 'black', '#F5F5F5')
    //     },
    //     SmallLabelStyle: {
    //         fontSize: "small",
    //         marginLeft: "5px",
    //         color: this.setStyles('white', 'white', 'black'),
    //         backgroundColor: this.setStyles('black', 'black', '#F5F5F5')
    //     },
    //     ImageStyle: {
    //         width: "55px",
    //         marginLeft: "1rem"
    //     }
    // });
    i18n.changeLanguage(detectBrowserLanguage());

    return (
        <I18nextProvider i18n={i18n} >
            <Dialog open={openTeamsDialog}>
                {/* <DialogTrigger disableButtonEnhancement>
                    <Button>Open dialog</Button>
                </DialogTrigger> */}
                <DialogSurface>
                    <DialogBody>
                        <DialogTitle>{teamsDialogTitle}</DialogTitle>
                        <DialogContent content={teamsDialogMessage}>
                            {/* <div dangerouslySetInnerHTML={{ __html: validationDialogMessage }} /> */}
                            {/* <TeamsDetails
                                graphClient={props.graphClient}
                                teamTitle={title}
                            /> */}
                        </DialogContent>
                        <DialogActions>
                            <DialogTrigger disableButtonEnhancement>
                                <Button appearance="secondary"
                                    onClick={() => {
                                        setOpenTeamsDialog(false);
                                    }}>Close</Button>
                            </DialogTrigger>
                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
            <Dialog open={openValidationDialog}>
                {/* <DialogTrigger disableButtonEnhancement>
                    <Button>Open dialog</Button>
                </DialogTrigger> */}
                <DialogSurface>
                    <DialogBody>
                        <DialogTitle>{validationDialogTitle}</DialogTitle>
                        <DialogContent>
                            <div dangerouslySetInnerHTML={{ __html: validationDialogMessage }} />
                        </DialogContent>
                        <DialogActions>
                            <DialogTrigger disableButtonEnhancement>
                                <Button appearance="secondary"
                                    onClick={() => {
                                        setOpenValidationDialog(false);
                                    }}>Close</Button>
                            </DialogTrigger>
                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
            <div className={useStackClassName()}>
                <Field>
                    <div style={{ display: 'flex', paddingTop: '1rem' }}>
                        <div>
                            <h3>{i18n.t('raiseRequest.header')} </h3>
                        </div>
                        <div style={{ width: "50px" }}>
                            <Image
                                shape="circular"
                                fit="contain"
                                src="https://img.icons8.com/cute-clipart/2x/help.png"
                            ></Image>
                        </div>
                    </div>
                    <div>
                        <text>{i18n.t('raiseRequest.sub-header-1')}</text>
                        <br />
                        <text>{i18n.t('raiseRequest.sub-header-2')}</text>
                    </div>
                </Field>
                <Field
                    label={i18n.t('raiseRequest.teamTitle')}
                    size="medium">
                    <div style={{ display: 'flex' }}>
                        <div >
                            <Input
                                style={{ width: "550px" }}
                                id="title"
                                placeholder={i18n.t('raiseRequest.teamTitle-placeholder')}
                                onChange={(event: ChangeEvent<HTMLInputElement>, data: InputOnChangeData) => {
                                    setTitle(data.value);
                                    setNameVerified(data.value === "" ? false : true);
                                }}
                            />
                        </div>
                        <div style={{ paddingLeft: "1rem" }}>
                            < Button
                                appearance="primary"
                                onClick={() => {
                                    setOpenTeamsDialog(true);
                                }}>
                                {i18n.t('raiseRequest.checkAvailability')}
                            </Button >
                        </div>
                    </div>
                </Field>
                <Field
                    label={i18n.t('raiseRequest.description')}
                    size="medium">
                    <Textarea
                        style={{ width: "550px" }}
                        id="description"
                        placeholder={i18n.t('raiseRequest.description-placeholder')}
                        onChange={(event: ChangeEvent<HTMLTextAreaElement>, data: InputOnChangeData) => {
                            setDescription(data.value);
                        }}
                    />
                </Field>
                <Field
                    label={i18n.t('raiseRequest.requester') + " " + i18n.t('raiseRequest.requester-desc')}
                    validationState={valid ? "none" : "warning"}
                    validationMessage={valid ? "" : "Input is limited to 20 characters."} >
                    <div style={{ display: 'flex' }}>
                        <div>
                            <User
                                ref={refRequester}
                                graphClient={props.graphClient}
                                title={"Requester"}
                                subTitle={"Requester"}
                                placeholder={"Who is the requester?"}
                                multiSelect={false}
                            />
                        </div>
                    </div>
                </Field>
                <Field
                    label={i18n.t('raiseRequest.owner') + " " + i18n.t('raiseRequest.owner-desc')}
                    size="medium">
                    <div style={{ display: 'flex' }}>
                        <User
                            ref={refOwner}
                            graphClient={props.graphClient}
                            title={"Owner"}
                            subTitle={"Team Owner"}
                            placeholder={"...and the owner"}
                            multiSelect={false}
                        />
                    </div>
                </Field>
                <Field
                    label={i18n.t('raiseRequest.coOwner') + " " + i18n.t('raiseRequest.coOwner-desc')}
                    size="medium">
                    <div style={{ display: 'flex' }}>
                        <User
                            ref={refCoOwners}
                            graphClient={props.graphClient}
                            title={"CoOwner"}
                            subTitle={"CoOwners"}
                            placeholder={"Finally the CoOwners pls"}
                            multiSelect={false}
                        />
                    </div>
                </Field>
                <Field
                    label={i18n.t('raiseRequest.businessJustification')}
                    size="medium">
                    <Textarea
                        id="justification"
                        style={{ width: "550px" }}
                        placeholder={i18n.t('raiseRequest.businessJustification-placeholder')}
                        onChange={(event: ChangeEvent<HTMLTextAreaElement>, data: InputOnChangeData) => {
                            setJustification(data.value);
                        }}
                    />
                </Field>
                <Field
                    label={i18n.t('raiseRequest.teamLifeTime')}
                    size="medium">
                    <Dropdown
                        id="lifetime"
                        style={{ width: "300px" }}
                        placeholder={i18n.t('raiseRequest.teamLifeTime-heading')}
                        onOptionSelect={(event, data) => {
                            setLifetime(data.optionValue as string)
                        }}
                    >
                        <Option >
                            {i18n.t('raiseRequest.teamLifeTime-option1')}
                        </Option>
                        <Option >
                            {i18n.t('raiseRequest.teamLifeTime-option2')}
                        </Option>
                        <Option >
                            {i18n.t('raiseRequest.teamLifeTime-option3')}
                        </Option>
                    </Dropdown>
                </Field>
                <Field
                    label={i18n.t('raiseRequest.privacyLevel')}
                    size="medium">
                    <Dropdown
                        id="privacy"
                        style={{ width: "300px" }}
                        placeholder={i18n.t('raiseRequest.privacyLevel-heading')}
                        onOptionSelect={(event, data) => {
                            setPrivacy(data.optionValue as string)
                        }}
                    >
                        <Option >
                            {i18n.t('raiseRequest.privacyLevel-option1')}
                        </Option>
                        <Option >
                            {i18n.t('raiseRequest.privacyLevel-option2')}
                        </Option>
                    </Dropdown>
                </Field>
                <Field
                    label={i18n.t('raiseRequest.sensitivityLevel')}
                    size="medium">
                    <Dropdown
                        id="sensitivity"
                        style={{ width: "300px" }}
                        placeholder={i18n.t('raiseRequest.sensitivityLevel-heading')}
                        onOptionSelect={(event, data) => {
                            setSensitivity(data.optionValue as string)
                        }}
                    >
                        <Option >{i18n.t('raiseRequest.sensitivityLevel-option1')}</Option>
                        <Option >{i18n.t('raiseRequest.sensitivityLevel-option2')}</Option>
                        <Option >{i18n.t('raiseRequest.sensitivityLevel-option3')}</Option>
                        <Option >{i18n.t('raiseRequest.sensitivityLevel-option4')}</Option>
                        <Option >{i18n.t('raiseRequest.sensitivityLevel-option5')}</Option>
                        <Option >{i18n.t('raiseRequest.sensitivityLevel-option6')}</Option>
                    </Dropdown>
                </Field>
                <Field
                    style={{ paddingTop: '10px', paddingBottom: '10px' }}>
                    <Button
                        id="btnSubmit"
                        style={{ width: "150px" }}
                        appearance="primary"
                        type="submit"
                        onClick={handleSubmitRequest}
                    >{i18n.t('raiseRequest.submitRequest')}</Button>
                </Field>
            </div>
        </I18nextProvider >
    );

}

export default (RaiseRequest);