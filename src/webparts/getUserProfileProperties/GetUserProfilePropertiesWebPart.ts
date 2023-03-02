import { sp } from "@pnp/sp";

import "@pnp/sp/profiles";
import {  
    Version  
} from '@microsoft/sp-core-library';  
import {  
    BaseClientSideWebPart,  
    
} from '@microsoft/sp-webpart-base';  
 //import { MSGraphClientV3 } from '@microsoft/sp-http';  
import {IPropertyPaneConfiguration,  
    PropertyPaneTextField  } from '@microsoft/sp-property-pane'
import styles from './components/GetUserProfileProperties.module.scss'; 
import * as strings from 'GetUserProfilePropertiesWebPartStrings';  
import {  
    IGetUserProfilePropertiesProps  
} from './components/IGetUserProfilePropertiesProps';
//import { ISPFXContext } from "@pnp/common/spfxcontextinterface";

export interface Istateset{
    state: { profiledata: any[]; }; 
}  
export default class GetUserProfilePropertiesWebPart extends BaseClientSideWebPart < IGetUserProfilePropertiesProps > {
    

    state: { profiledata: any[]; }; 
    constructor(props:Istateset){
        super()
        this.state={
            profiledata:[],
            
        }   
        this.GetUserProperties.bind(this);
    }
    

    private GetUserProperties(): void{  

        sp.profiles.getPropertiesFor("i:0#.f|membership|JoniS@8bh87k.onmicrosoft.com").then (function(result1) 
        {  
            var userProperties1 = result1;
            console.log( userProperties1)  ;
    }).catch(function(error: string) {  
        console.log("Error: " + error);  
        });  
    
        // sp.profiles.get().then (function(result1) 
        // {  
        //     var userProperties1 = result1.UserProfileProperties1;
        //     console.log( userProperties1)  ;
        //     var userPropertyValues1 = "";  

        //     userProperties1.map((x:any)=> {
        //         userPropertyValues1 +=  userProperties1[0].Key + " - " +  userProperties1[0].Value + "<br/>";
        //         console.log(userPropertyValues1);
        //        // document.getElementById("spUserProfileProperties").innerHTML=userPropertyValues;
        //         this.domElement.getElementById("spUserProfileProperties").innerHTML=  userPropertyValues1; 
        //         // this.document.getElementById("spUserProfileProperties").innerHTML=userPropertyValues;
                
        //     })
        // }).catch(function(error: string) {  
        //     console.log("Error: " + error);  
        //     }); 
         //for all users
         
        //  this.context.msGraphClientFactory 
        //  .getClient('3')  
        //  .then((client: MSGraphClientV3): void => {  
        //    client.api('/users').get(async (error:any, users: any, rawResponse?: any) => {  
        //        //console.log(users.value);  
        //        for(let user of users.value){  
        //          console.log(user.userPrincipalName);  
        //          let loginName="i:0#.f|membership|"+user.userPrincipalName;  
        //          const profile = await sp.profiles.getPropertiesFor(loginName);  
        //          console.log(profile);   
        //        }  
        //    });  
        //  }); 




       
            // var userPropertyValues1 = "";  

            // userProperties1.map((x:any)=> {
            //     userPropertyValues1 +=  userProperties1[0].Key + " - " +  userProperties1[0].Value + "<br/>";
            //     console.log(userPropertyValues1);
            //    // document.getElementById("spUserProfileProperties").innerHTML=userPropertyValues;
            //     this.domElement.getElementById("spUserProfileProperties").innerHTML=  userPropertyValues1; 
            //     // this.document.getElementById("spUserProfileProperties").innerHTML=userPropertyValues;
                
            // })
           
    
        sp.profiles.myProperties.get().then(function(result:any) 
        {  
            var userProperties = result.UserProfileProperties;
            console.log( userProperties)  ;
            var userPropertyValues = "";  

            userProperties.map((x:any)=> {
                userPropertyValues +=  userProperties[0].Key + " - " +  userProperties[0].Value + "<br/>";
                console.log(userPropertyValues);
               // document.getElementById("spUserProfileProperties").innerHTML=userPropertyValues;
                this.domElement.getElementById("spUserProfileProperties").innerHTML=  userPropertyValues; 
                // this.document.getElementById("spUserProfileProperties").innerHTML=userPropertyValues;
                
            })
           
        //     forEach(function( userProperties) {  
        //         userPropertyValues +=  userProperties.Key + " - " +  userProperties.Value + "<br/>";  
        //     });  
        //     console.log( userProperties, userPropertyValues)
        //   this.document.getElementById("spUserProfileProperties").innerHTML=userPropertyValues;
        // //    this.domElement.getElementById("spUserProfileProperties")=  userPropertyValues; 
         }).catch(function(error: string) {  
        console.log("Error: " + error);  
        });  
    }  
    
    

  
    public render(): void {  
        this.domElement.innerHTML = `  
<div class="${styles.helloWorld}">  
<div class="${styles.container}">  
<div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
<div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPusstrong ms-u-lgPusstrong">  
<span class="ms-font-xl ms-fontColor-white" style="font-size:28px">Welcome to SharePoint Framework Development using PnP JS Library</span>  
<p class="ms-font-l ms-fontColor-white" style="text-align: left">Demo : Retrieve User Profile Properties</p>  
</div>  
</div>  
<div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
<div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">User Profile Details</div>  
<br>  
<div id="spUserProfileProperties" />  
</div>  
</div>  
</div>`;  
        this.GetUserProperties();  
    }  
    protected async onInit(): Promise<void> {  
  
        await super.onInit();  
        
        // spfi().using(SPFx({pageContext: this.context.pageContext}));
        // // other init code may be present  
        
    //    sp.setup(spfxContext:this.context);  
      }
    protected get dataVersion(): Version {  
        return Version.parse('1.0');  
    }  
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {  
        return {  
            pages: [{  
                header: {  
                    description: strings.PropertyPaneDescription  
                },  
                groups: [{  
                    groupName: strings.BasicGroupName,  
                    groupFields: [  
                        PropertyPaneTextField('description', {  
                            label: strings.DescriptionFieldLabel  
                        })  
                    ]  
                }]  
            }]  
        };  
    }  
}  




// function spfi() {
//     throw new Error("Function not implemented.");
// }

// function SPFx(arg0: ISPFXContext): any {
//     throw new Error("Function not implemented.");}
// function forEach(arg0: (property: any) => void) {
//   throw new Error('Function not implemented.');
// }
// function getElementById(arg0: void) {
//   throw new Error('Function not implemented.');
// }

