import { Option, Combobox, useId, Input, ComboboxProps, Button, Label } from "@fluentui/react-components";
import { BoxSearch24Regular, Dismiss12Regular, Search24Regular, Search32Regular } from "@fluentui/react-icons";
import { Client } from "@microsoft/microsoft-graph-client";
import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from "react";
import { UserDetails } from "./UserDetails";
import { useStyles } from "./Dummy";


const User = forwardRef(function User(props: {
    graphClient: Client;
    title: string,
    subTitle: string,
    placeholder: string,
    multiSelect: boolean
}, ref: any) {
    const styles = useStyles();

    // generate ids for handling labelling
    const comboId = useId("cmbRequester");
    const selectedListId = `${comboId}-selection`;

    // refs for managing focus when removing tags
    const selectedListRef = useRef<HTMLUListElement>(null);
    const comboboxInputRef = useRef<HTMLInputElement>(null);

    const [users, setUsers] = useState([]);
    const [selectedUsers, setSelectedUsers] = useState<string[]>([]);
    const [search, setSearch] = useState("");

    const searchUser = () => {
        console.log("comboboxInputRef", comboboxInputRef)
        let url = " ";
        if (search) {
            url = `https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${search}')`;
        } else {
            url = `https://graph.microsoft.com/v1.0/users`;
        }
        props.graphClient.api(url)
            .select("userPrincipalName, displayName, jobTitle")
            .top(10)
            .get().then((response) => {
                console.log("response", response);
                setUsers(response.value);
            });
    }

    const onOptionSelect: ComboboxProps["onOptionSelect"] = (event, data) => {
        console.log("data.optionText", data.optionText);
        console.log("data.optionValue", data.optionValue);

        let foo: string[] = []
        data.selectedOptions.map((o) => {
            console.log("o", o);
            foo.push(o)
        })
        console.log("foo", foo);
        console.log("data.selectedOptions", data.selectedOptions)
        // foo.push(data.optionValue)
        // setSelectedUsers(foo);
        setSelectedUsers(data.selectedOptions);
        // setSelectedUsers(selectedUsers.filter((o) => o !== data.optionText));
    };

    const onTagClick = (option: string, index: number) => {
        // remove selected option
        setSelectedUsers(selectedUsers.filter((o) => o !== option));

        // focus previous or next option, defaulting to focusing back to the combo input
        const indexToFocus = index === 0 ? 1 : index - 1;
        const optionToFocus = selectedListRef.current?.querySelector(
            `#${comboId}-selected-${indexToFocus}`
        );
        if (optionToFocus) {
            (optionToFocus as HTMLButtonElement).focus();
        } else {
            comboboxInputRef.current?.focus();
        }
    };

    const labelledBy = selectedUsers.length > 0 ? `${comboId} ${selectedListId}` : comboId;

    useImperativeHandle(ref, () => {
        return {
            getSelectedUsers() {
                return selectedUsers;
            }
        }
    }, [selectedUsers]);

    return (
        <div>
            {/* <label id={comboId}>{props.title}</label> */}
            <Combobox
                // style={{ marginLeft: '5px' }}
                multiselect={props.multiSelect}
                placeholder={props.placeholder}
                selectedOptions={selectedUsers}
                ref={comboboxInputRef}
                onChange={(event) => {
                    console.log("event", event, event.target.value);
                    setSearch(event.target.value);
                }}
                onOptionSelect={onOptionSelect}
            >
                {users.map((item: any) => (
                    <Option
                        key={item.userPrincipalName}
                        text={item.displayName}
                        value={item.userPrincipalName}>
                        <UserDetails
                            graphClient={props.graphClient}
                            userPrincipalName={item.userPrincipalName} />
                    </Option>
                ))}
            </Combobox>
            <label>
                <Search32Regular
                    style={{ verticalAlign: 'middle' }}
                    onClick={searchUser} />
            </label>
            {selectedUsers.length ? (
                <ul
                    id={selectedListId}
                    className={styles.tagsList}
                    ref={selectedListRef}
                >
                    <span
                        id={`${comboId}-selected`}
                        style={{ marginBottom: '5px' }}>
                        {props.subTitle}:
                    </span>
                    {selectedUsers.map((option, i) => (
                        <li key={option}>
                            <Button
                                size="small"
                                shape="circular"
                                appearance="primary"
                                icon={<Dismiss12Regular />}
                                iconPosition="after"
                                onClick={() => onTagClick(option, i)}
                                id={`${comboId}-selected-${i}`}
                                aria-labelledby={`${comboId}-selected ${comboId}-selected-${i}`}
                            >
                                {option}
                            </Button>
                        </li>
                    ))}
                </ul>
            ) : null}
        </div>
    );
}
)
export default User;