import { TeamsUserCredentialAuthConfig, TeamsUserCredential } from "@microsoft/teamsfx";
import { useEffect, useState } from "react";
import { API_SCOPES, LIST_ID, SITE_ID } from "../Constants";
import { useGraphWithCredential, useTeamsFx, useTeamsUserCredential } from "@microsoft/teamsfx-react";
import { Client } from "@microsoft/microsoft-graph-client";
import { Login, PeoplePicker, Person, Providers, useIsSignedIn } from "@microsoft/mgt-react";
import { registerMgtComponents } from "@microsoft/mgt-components";
import { Msal2Provider, Msal2Config, Msal2PublicClientApplicationConfig } from '@microsoft/mgt-msal2-provider';
import { Button, Persona, Table, TableBody, TableCell, TableCellActions, TableCellLayout, TableHeader, TableHeaderCell, TableRow, Tooltip } from "@fluentui/react-components";
import { AppTitleRegular, EditRegular, Person24Filled, Person24Regular, UsbPlugRegular } from "@fluentui/react-icons";
import { getRequests, getUserDisplayName } from "../services/GraphAPIs";
import { Link, useNavigate } from "react-router-dom";
import { pages } from "@microsoft/teams-js";


// Providers.globalProvider = new Msal2Provider({ clientId: '0be5fed3-eaaa-4ce8-a04a-957d94c18b47' });
// registerMgtComponents();
// const [isSignedIn] = useIsSignedIn();

function UserDetails(props: { graphClient: Client, userPrincipalName: string }) {
    const [displayName, setDisplayName] = useState();
    const [photo, setPhoto] = useState(Object);
    const [presence, setPresence] = useState(Object);
    const [jobTitle, setJobTitle] = useState();
    const [userId, setUserId] = useState();

    useEffect(() => {
        // User Id, DsiplayName, JobTitle
        props.graphClient.api(`/users/${props.userPrincipalName}`)
            .select("id, displayName, jobTitle")
            .get()
            .then((response: any) => {
                // console.log("setDisplayName..", response.displayName, response);
                setDisplayName(response.displayName);
                setJobTitle(response.jobTitle);
                setUserId(response.id);
            }).then(() => {
                // User Presence
                props.graphClient.api(`/users/${userId}/presence`)
                    .get()
                    .then((response: any) => {
                        console.log("setPresence..", response);
                        setPresence((response.availability as string).toLowerCase());
                        // setPresence("Available")
                    });
            });

        // User Photo
        props.graphClient.api(`/users/${props.userPrincipalName}/photos/64x64/$value`)
            .get()
            .then((response: any) => {
                // console.log("setPhoto..", response);
                setPhoto(URL.createObjectURL(response));
                // console.log("setPhoto..", URL.createObjectURL(response));

            })
            .catch((reason: any) => {
                // console.log("image not found for: ", props.userPrincipalName)
            });
    }, [])

    // let url: any = window.URL + '' + window.webkitURL;
    // var blobUrl = url.createObjectURL(photo);
    return (
        <div>
            <Persona
                name={displayName}
                secondaryText={jobTitle}
                presence={{ status: presence }}
                size="extra-large"
                avatar={{
                    image: {
                        src: photo,
                    },
                }}
            />
        </div>
    );
}

export function MyRequests(props: { graphClient: Client }) {

    const [requests, setRequests] = useState([]);
    const navigate = useNavigate();

    const columns = [
        { columnKey: "title", label: "Title" },
        { columnKey: "description", label: "Description" },
        { columnKey: "requester", label: "Requester" },
        { columnKey: "owner", label: "Owner" },
        { columnKey: "coOwner", label: "CoOwner" },
        { columnKey: "lifetime", label: "Lifetime" },
    ];

    useEffect(() => {

        // ///replaced in the GraphAPIs.ts. Testing still required.
        // props.graphClient.api(`/sites/${SITE_ID}/lists/${LIST_ID}/items`)
        //     // .version("beta")
        //     .expand("fields")
        //     // .filter(`fields/Requester eq '${userPrincipalName}'`)
        //     // .top(5)
        //     .get()
        //     .then((response: any) => {
        //         setRequests(response.value);
        //         // console.log("requests..", response.value);

        //         for (let ctr = 0; ctr < response.value.length; ctr++) {
        //             const request = response.value[ctr];
        //             // console.log("request.fields.Title", request.fields.Title)
        //         }
        //     })

        getRequests(props.graphClient, "")
            .then((reqs) => {
                setRequests(reqs);
            })

    }, [])

    function trimText(strInput: string) {
        return strInput.length > 25 ? strInput.substring(0, 25) + "..." : strInput
    }

    return (
        <div style={{ padding: '10px', margin: '10px' }}>
            <Table arial-label="Default table">
                <TableHeader>
                    <TableRow >
                        {columns.map((column) => (
                            <TableHeaderCell key={column.columnKey}>
                                {column.label}
                            </TableHeaderCell>
                        ))}
                    </TableRow>
                </TableHeader>
                <TableBody>
                    {requests.map((item: any) => (
                        <TableRow key={item.fields.id}>
                            <TableCell >
                                <Tooltip content={item.fields.Title} relationship="label" >
                                    <TableCellLayout media={<AppTitleRegular />}>
                                        {trimText(item.fields.Title)}
                                    </TableCellLayout>
                                </Tooltip>
                            </TableCell>
                            <TableCell>
                                <Tooltip content={item.fields.Description.replace(/(<([^>]+)>)/ig, '')} relationship="label" >
                                    <TableCellLayout>
                                        {trimText(item.fields.Description.replace(/(<([^>]+)>)/ig, ''))}
                                    </TableCellLayout>
                                </Tooltip>
                            </TableCell>
                            <TableCell>
                                <Tooltip content={item.fields.Requester} relationship="label" >
                                    <TableCellLayout>
                                        <UserDetails
                                            graphClient={props.graphClient}
                                            userPrincipalName={item.fields.Requester}
                                        />
                                    </TableCellLayout>
                                </Tooltip>
                            </TableCell>
                            <TableCell>
                                <Tooltip content={item.fields.Owner} relationship="label" >
                                    <TableCellLayout>
                                        <UserDetails
                                            graphClient={props.graphClient}
                                            userPrincipalName={item.fields.Owner}
                                        />
                                    </TableCellLayout>
                                </Tooltip>
                            </TableCell>
                            <TableCell>
                                <Tooltip content={item.fields.CoOwner} relationship="label" >
                                    <TableCellLayout>
                                        <UserDetails
                                            graphClient={props.graphClient}
                                            userPrincipalName={item.fields.CoOwner}
                                        />
                                    </TableCellLayout>
                                </Tooltip>
                            </TableCell>
                            <TableCell>
                                <TableCellLayout>
                                    {item.fields.Lifetime}
                                </TableCellLayout>
                            </TableCell>
                            <TableCell>
                                <TableCellLayout>

                                </TableCellLayout>
                            </TableCell>
                            <TableCell>
                                <TableCellLayout>
                                    <Link to={"/raiseRequest"} >
                                        <Button
                                            icon={<EditRegular />}
                                            onClick={() => {

                                                pages.currentApp
                                                    .navigateTo({ pageId: "raiseRequest" })
                                                    .then(() => {
                                                        navigate("/raiseRequest", { replace: true, state: { id: 1 } });
                                                    });
                                            }} />
                                    </Link>
                                </TableCellLayout>
                            </TableCell>
                        </TableRow>
                    ))}
                </TableBody>
            </Table>
        </div>
    );
}