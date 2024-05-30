import { Persona } from "@fluentui/react-components";
import { Client } from "@microsoft/microsoft-graph-client";
import { useEffect, useState } from "react";

export function UserDetails(props: { graphClient: Client; userPrincipalName: string; }) {
    const [displayName, setDisplayName] = useState();
    const [photo, setPhoto] = useState(Object);
    const [presence, setPresence] = useState(Object);
    const [jobTitle, setJobTitle] = useState();
    const [userId, setUserId] = useState();

    useEffect(() => {
        // User Id, DsiplayName, JobTitle
        props.graphClient
            .api(`/users/${props.userPrincipalName}`)
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
                        // console.log("setPresence..", response);
                        setPresence((response.availability as string).toLowerCase());
                        // setPresence("Available")
                    });
            });
        // User Photo
        props.graphClient
            .api(`/users/${props.userPrincipalName}/photos/64x64/$value`)
            .get()
            .then((response: any) => {
                // console.log("setPhoto..", response);
                setPhoto(URL.createObjectURL(response));
                // console.log("setPhoto..", URL.createObjectURL(response));
            })
            .catch((reason: any) => {
                // console.log("image not found for: ", props.userPrincipalName)
            });
    }, []);

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
                }} />
        </div>
    );
}
