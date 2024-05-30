import { Table, TableBody, TableCell, TableCellLayout, TableRow, Tooltip, shorthands, tokens, makeStyles, Input, useId, Button } from "@fluentui/react-components";
import { AppTitleRegular } from "@fluentui/react-icons";
import { Client } from "@microsoft/microsoft-graph-client";
import { useRef, useState } from "react";
import User from "./User";

export const useStyles = makeStyles({
    root: {
        // Stack the label above the field with a gap
        display: "grid",
        gridTemplateRows: "repeat(1fr)",
        justifyItems: "start",
        ...shorthands.gap("2px"),
        maxWidth: "400px",
    },
    tagsList: {
        listStyleType: "none",
        marginBottom: tokens.spacingVerticalXXS,
        marginTop: 0,
        paddingLeft: 0,
        display: "flex",
        gridGap: tokens.spacingHorizontalXXS,
    },
});

export function Dummy(props: { graphClient: Client }) {
    const searchId = useId("txtSearch");

    const [searchText, setSearchText] = useState("a");

    const refRequester: any = useRef();
    const refOwner: any = useRef();
    const refCoOwners: any = useRef();


    return (
        <div style={{ padding: '20px' }}>
            <User
                ref={refRequester}
                graphClient={props.graphClient}
                title={"Requester"}
                subTitle={"Requester"}
                placeholder={"Who is the requester?"}
                multiSelect={false}
            />
            <br /><br />
            <User
                ref={refOwner}
                graphClient={props.graphClient}
                title={"Owner"}
                subTitle={"Team Owner"}
                placeholder={"...and the owner"}
                multiSelect={false}
            />
            <br /><br />
            <User
                ref={refCoOwners}
                graphClient={props.graphClient}
                title={"CoOwner"}
                subTitle={"CoOwners"}
                placeholder={"Finally the CoOwners pls"}
                multiSelect={true}
            />
            <Button onClick={() => {
                console.log("refRequester", refRequester.current);
                var foo = refRequester.current.getSelectedUsers();
                console.log("foo", foo);
            }} />
        </div>
    );
}

