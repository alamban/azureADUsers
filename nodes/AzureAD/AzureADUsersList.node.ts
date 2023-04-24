import { IExecuteFunctions } from 'n8n-core';
import { 
    INodeType, 
    INodeTypeDescription, 
    INodeExecutionData, 
    NodeOperationError 
} from 'n8n-workflow';

export class AzureADUsersList implements INodeType {
	description: INodeTypeDescription = {
        displayName: 'Azure AD Users List',
        name: 'AzureADUsersList',
        icon: 'file:azuread.png',
        group: ['transform'],
        version: 1,
        subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
        description: 'Get data from NASAs API',
        defaults: {
        	name: 'Azure AD Users List',
        },
        inputs: ['main'],
        outputs: ['main'],
        credentials: [
        	{
        		name: 'AzureADApi',
        		required: true,
        	},
        ],

		// Basic node details will go here
		properties: [
            {
                displayName: 'Information to fetch',
                name: 'infoToFetch',
                type: 'options',
                noDataExpression: true,
                options: [
                    {
                        name: 'Users',
                        value: 'users',
                    },
                ],
                default: 'Users',
            },
        
        ]
        
	};


    async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		var items = this.getInputData();

		let item: INodeExecutionData;
		let myString: string;

        let httpBasicAuth;
        httpBasicAuth = await this.getCredentials('AzureADApi');
        console.log('Creds', httpBasicAuth);

		// Iterates over all input items and add the key "myString" with the
		// value the parameter "myString" resolves to.
		// (This could be a different value for each item in case it contains an expression)
		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				myString = this.getNodeParameter('AzureADApi', itemIndex, '') as string;
				item = items[itemIndex];

				item.json['AzureADApi'] = myString;
				myString = this.getNodeParameter('infoToFetch', itemIndex, '') as string;
				item = items[itemIndex];

				item.json['infoToFetch'] = myString;

				// myString = this.getNodeParameter('tenantId', itemIndex, '') as string;
				// item = items[itemIndex];

				// item.json['tenantId'] = myString;

				// myString = this.getNodeParameter('clientId', itemIndex, '') as string;
				// item = items[itemIndex];

				// item.json['clientId'] = myString;

				// myString = this.getNodeParameter('clientSecret', itemIndex, '') as string;
				// item = items[itemIndex];

				// item.json['clientSecret'] = myString;
				// console.log(items);
				// console.log(items[0].json.clientId);

				// ---------------------------------------------- USER CATCH ----------------------------------------------------

				const {
					Client
				} = require("@microsoft/microsoft-graph-client");
				const {
					TokenCredentialAuthenticationProvider
				} = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
				const {
					ClientSecretCredential
				} = require("@azure/identity");
				let tenantId = `${httpBasicAuth.tenantId}`// "a755254c-9686-4300-a8ef-525b712e12c6"
				let clientId = `${httpBasicAuth.clientId}`//"6912bed7-46e4-427c-a8e4-66c7fdc6d854"
				let clientSecret = `${httpBasicAuth.clientSecret}`// "GAJ8Q~4AXUtrLn5cV8cDuatpDSzBvHEu9Lba0cyy"
				const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
				const authProvider = new TokenCredentialAuthenticationProvider(credential, {
					scopes: ["https://graph.microsoft.com/.default"]
				});
				const client = Client.initWithMiddleware({
					debugLogging: true,
					authProvider
					// Use the authProvider object to create the class.
				});

				let users = await client.api(`/${items[0].json.infoToFetch}`).get();
				console.log('USERS:: -->    ', users);
				console.log('USERS Count:: -->  ', users.value.length);
				// var userNames: string[] | undefined = undefined;
				// users.value.map((elem : any) => {
				//     userNames.push(elem.displayName);
				// 	return null;
				// });
				let finData = [{
					json: {
						users: {}
					}
				}];
				finData[0].json.users = users.value;
				items = finData;
				// for (let user = 0; user < users.value.length; user++) {
				// 	userNames.push()
				// }
				// console.log('USERNAMES:: -->    ', userNames);

			} catch (error) {
				// This node should never fail but we want to showcase how
				// to handle errors.
				if (this.continueOnFail()) {
					items.push({ json: this.getInputData(itemIndex)[0].json, error, pairedItem: itemIndex });
				} else {
					// Adding `itemIndex` allows other workflows to handle this error
					if (error.context) {
						// If the error thrown already contains the context property,
						// only append the itemIndex
						error.context.itemIndex = itemIndex;
						throw error;
					}
					throw new NodeOperationError(this.getNode(), error, {
						itemIndex,
					});
				}
			}
		}

		return this.prepareOutputData(items);
	}
}
