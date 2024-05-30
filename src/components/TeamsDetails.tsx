export default function TeamsDetails(props: { graphClient: any, teamTitle: string }) {
return(<div></div>)

}

// import { useEffect, useState } from "react";
// import { getTeamDetails } from "../services/GraphAPIs";
// import { Card, CardHeader, Input } from "@fluentui/react-components";
// import i18n from "../i18n";

// export default function TeamsDetails(props: { graphClient: any, teamTitle: string }) {

//     const [teams, setTeams] = useState<any[]>([]);
//     const [orgTeamsList, setOrgTeamsList] = useState<any[]>([]);

//     useEffect(() => {

//         async function teamDetails() {
//             const results = await getTeamDetails(props.graphClient, props.teamTitle);
//             console.log("fetching data...")
//             setTeams(results.value);
//             setOrgTeamsList(results.value);
//         }

//         teamDetails();

//     }, []);


//     if (orgTeamsList != null) {
//         console.log("teams.length", orgTeamsList.length);
//         if (orgTeamsList.length === 0) {
//             return (<h3>{i18n.t('raiseRequest.teamDetails-2')}</h3>);
//         }
//         return (
//             <div>
//                 <div style={{ margin: '10px 0px 10px 0px' }}>
//                     <Input
//                         style={{}}
//                         // icon={<SearchIcon />}
//                         placeholder={i18n.t('raiseRequest.teamDetailsSearch.placeholder')}
//                         onChange={(event: any) => {
//                             let value = event.currentTarget.value;
//                             let searchValue = new RegExp(value, 'i');

//                             console.log("value", value.length);
//                             let filteredTeams = orgTeamsList.filter((item: any) => {
//                                 console.log("item", item);
//                                 let filteredOwners = item.owners.filter((owner: any) => {
//                                     if (null != owner.displayName && owner.displayName.search(searchValue) > -1)
//                                         return true;
//                                     else
//                                         return false;
//                                 })
//                                 console.log("filteredOwners", filteredOwners);
//                                 if ((null != item.description && item.description.search(searchValue) > -1)
//                                     || (null != item.displayName && item.displayName.search(searchValue) > -1)
//                                     || (null != item.owners && filteredOwners.length > 0)
//                                 )
//                                     return true;
//                                 else
//                                     return false;
//                             });

//                             setTeams(filteredTeams);
//                         }}
//                     />
//                     <Text
//                         style={{ marginLeft: '10px' }}
//                         size='small'
//                         content={i18n.t('raiseRequest.teamDetailsSearch.description')}
//                     />
//                 </div>
//                 <div>
//                     {
//                         teams.map((team: any, index: any) => (
//                             <div key={index}>
//                                 <Card styles={{ blockSize: "auto", boxShadow: "1px 1px 1px 1px grey" }} fluid={true}>
//                                     <CardHeader>
//                                         <Flex>
//                                             <strong>{i18n.t('raiseRequest.teamDetailsSearchResult.Title')}&nbsp; </strong>
//                                             {team.displayName}
//                                         </Flex>
//                                     </CardHeader>
//                                     <CardBody fitted={true} key={team.id}>
//                                         <Flex column>
//                                             <TeamOwners owners={team.owners} />
//                                             <Flex>
//                                                 <strong>{i18n.t('raiseRequest.teamDetailsSearchResult.Type')}&nbsp;  </strong>
//                                                 {team.groupTypes}
//                                             </Flex>
//                                             <Flex>
//                                                 <strong>{i18n.t('raiseRequest.teamDetailsSearchResult.Visibility')}&nbsp;  </strong>
//                                                 {team.visibility}
//                                             </Flex>
//                                             <Flex>
//                                                 <strong>{i18n.t('raiseRequest.teamDetailsSearchResult.Description')}&nbsp;  </strong>
//                                                 {team.description}
//                                             </Flex>
//                                         </Flex>
//                                     </CardBody>
//                                 </Card>
//                                 <div style={{ height: "10px" }} ></div>
//                             </div>
//                         ))
//                     }
//                 </div>
//             </div>
//         );
//     } else {
//         return (<div></div>);
//     }

// }
