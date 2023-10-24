import { TeamsUserCredential } from "@microsoft/teamsfx";
import { createContext } from "react";
import { Theme } from "@fluentui/react-components";

export const TeamsFxContext = createContext<{
  theme?: Theme;
  themeString: string;
  teamsUserCredential?: TeamsUserCredential;
}>({
  theme: undefined,
  themeString: "",
  teamsUserCredential: undefined,
});

// var old = TeamsUserCredential.prototype.getToken;
// TeamsUserCredential.prototype.getToken = async (scopes: string | string[], options?: any) => {
//   // hook before call
//   let scopesArray = typeof scopes === "string" ? scopes.split(" ") : scopes;
//   scopesArray = scopesArray.map((scope) => {
//     if (scope.includes("https://graph.microsoft.com")) {
//       return "https://graph.microsoft.com/.default";
//     }
//     else {
//       return scope;
//     }
//   })

//   return await old(scopesArray, options);
//   // hook after call
// };