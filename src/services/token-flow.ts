import { window } from "vscode";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { GraphResult } from "../types/graph-result";

export class TokenFlowService extends ServiceBase {
	// The constructor for the TokenFlowService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
	}

	async enableHybridFlow(item: AppRegItem, enable: Boolean): Promise<void> {
        // Following the behavior in the portal, if Hybrid Flow is enabled and is being disabled, warn the user that this will disable it for the entire application and check if they want to proceed.
        if (item?.contextValue === "HYBRID-FLOW-ENABLED" && enable !== true){
            window
                .showWarningMessage(
                    "Turning off hybrid grant settings will disable them for the entire application. If you have an application in production still using hybbrid settings, disabling these settings may cause issues. Are you sure you want to disable hybrid grant?",
                    { modal: true }, 
                    "Disable"
                    )
                .then(async (selection) => {
                    if (selection === "Disable") {
                        await this.enableHybridFlowInternal(item, enable);
                    }
                });
        }
        else{
            await this.enableHybridFlowInternal(item, enable);
        }
    }
    
    private async enableHybridFlowInternal(item: AppRegItem, enable: Boolean): Promise<void> {
        const status = this.indicateChange("Updating Token Flow...", item);
        const update: GraphResult<void> = await this.graphRepository.updateApplication(item.objectId!, { web: {implicitGrantSettings: {enableIdTokenIssuance: enable === true } } });
        update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error);
    }

	async enableImplicitFlow(item: AppRegItem, enable: Boolean): Promise<void> {
        // Following the behavior in the portal, if Implicit Flow is enabled and is being disabled, warn the user that this will disable it for the entire application and check if they want to proceed.
        if (item?.contextValue === "IMPLICIT-FLOW-ENABLED" && enable !== true){
            window
                .showWarningMessage(
                    "Turning off implicit grant settings will disable them for the entire application. If you have an application in production still using implicit settings, disabling these settings may cause issues. Are you sure you want to disable implicit grant?",
                    { modal: true }, 
                    "Disable"
                    )
                .then(async (selection) => {
                    if (selection === "Disable") {
                        await this.enableImplicitFlowInternal(item, enable);
                    }
                });
        }
        else{
            await this.enableImplicitFlowInternal(item, enable);
        }
    }

    private async enableImplicitFlowInternal(item: AppRegItem, enable: Boolean): Promise<void> {
        const status = this.indicateChange("Updating Token Flow...", item);
        const update: GraphResult<void> = await this.graphRepository.updateApplication(item.objectId!, { web: {implicitGrantSettings: {enableIdTokenIssuance: enable === true, enableAccessTokenIssuance: enable === true } } });
        update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error);
    }

    async enablPublicClientFlows(item: AppRegItem, enable: Boolean): Promise<void> {
        const status = this.indicateChange("Updating Token Flow...", item);
        const update: GraphResult<void> = await this.graphRepository.updateApplication(item.objectId!, { isFallbackPublicClient: enable === true });
        update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error);
    }
}
