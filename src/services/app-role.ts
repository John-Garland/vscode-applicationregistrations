import { window } from "vscode";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { AppRole, Application } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from "uuid";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";
import { clearStatusBarMessage, setStatusBarMessage } from "../utils/status-bar";
import { validateAppRoleDescription, validateAppRoleDisplayName, validateAppRoleValue } from "../utils/validation";

export class AppRoleService extends ServiceBase {
	// The constructor for the AppRolesService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
	}

	// Adds a new app role to an application registration.
	async add(item: AppRegItem): Promise<void> {
		// Show that we're doing something
		const check = setStatusBarMessage("Reading App Roles...");

		// Get the existing app roles.
		const roles = await this.getAppRoles(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (roles === undefined) {
			return;
		}

		clearStatusBarMessage(check!);

		// Capture the new app role details by passing in an empty app role.
		const role = await this.inputRoleDetails({}, false, roles);

		// If the user cancels the input then return undefined.
		if (role === undefined) {
			return;
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Adding App Role...", item);

		// Add the new app role to the array.
		roles.push(role);

		// Update the application.
		await this.updateApplication(item.objectId!, { appRoles: roles }, status, "APP-ROLES");
	}

	// Edits an app role from an application registration.
	async edit(item: AppRegItem): Promise<void> {
		// Show that we're doing something
		const check = setStatusBarMessage("Reading App Role...");

		// Read the existing app roles.
		const roles = await this.getAppRoles(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (roles === undefined) {
			return;
		}

		clearStatusBarMessage(check!);

		// Capture the new app role details by passing in the existing role.
		const role = await this.inputRoleDetails(roles.filter((r) => r.id === item.value!)[0], true, roles);

		// If the user cancels the input then return undefined.
		if (role === undefined) {
			return;
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Updating App Role...", item);

		// Update the application.
		await this.updateApplication(item.objectId!, { appRoles: roles }, status, "APP-ROLES");
	}

	// Edits an app role value from an application registration.
	async editValue(item: AppRegItem): Promise<void> {
		// Show that we're doing something
		const check = setStatusBarMessage("Reading App Role...");

		// Read the existing app roles.
		const roles = await this.getAppRoles(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (roles === undefined) {
			return;
		}

		clearStatusBarMessage(check!);

		// Get the existing role.
		const role = roles.filter((r) => r.id === item.value!)[0];

		switch (item.contextValue) {
			case "ROLE-ENABLED":
			case "ROLE-DISABLED":
				// Prompt the user for the new display name.
				const displayName = await this.inputDisplayName("Edit Display Name (1/1)", role.displayName!, validateAppRoleDisplayName);

				// If escape is pressed then return undefined.
				if (displayName === undefined) {
					return undefined;
				}

				role.displayName = displayName;
				break;
			case "ROLE-VALUE":
				// Prompt the user for the new value.
				const value = await this.inputValue("Edit App Role Value (1/1)", true, roles, role.value!, validateAppRoleValue);

				// If escape is pressed then return undefined.
				if (value === undefined) {
					return undefined;
				}

				role.value = value;
				break;
			case "ROLE-DESCRIPTION":
				// Prompt the user for the new display name.
				const description = await this.inputDescription("Edit Description (1/1)", role.description!, validateAppRoleDescription);

				// If escape is pressed then return undefined.
				if (description === undefined) {
					return undefined;
				}

				role.description = description;
				break;
			case "ROLE-ALLOWED":
				// Prompt the user for the new allowed member types.
				const allowed = await this.inputAllowedMemberTypes("Edit App Role (1/1)");

				// If escape is pressed or the new allowed member types is empty then return undefined.
				if (allowed === undefined) {
					return undefined;
				}

				role.allowedMemberTypes = allowed.value;
				break;
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Updating App Role...", item);

		// Update the application.
		await this.updateApplication(item.objectId!, { appRoles: roles }, status, "APP-ROLES");
	}

	// Changes the enabled state of an app role from an application registration.
	async changeState(item: AppRegItem, state: boolean): Promise<void> {
		// Set the added trigger to the status bar message.
		const status = this.indicateChange(state === true ? "Enabling App Role..." : "Disabling App Role...", item);

		// Get the parent application so we can read the app roles.
		const roles = await this.getAppRoles(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (roles === undefined) {
			return;
		}

		// Toggle the state of the app role.
		roles.filter((r) => r.id === item.value!)[0].isEnabled = state;

		// Update the application.
		await this.updateApplication(item.objectId!, { appRoles: roles }, status, "APP-ROLES");
	}

	// Deletes an app role from an application registration.
	async delete(item: AppRegItem, hideConfirmation: boolean = false): Promise<void> {
		if (item.state !== false && hideConfirmation === false) {
			const disableRole = await window.showWarningMessage(`The App Role ${item.label} cannot be deleted unless it is disabled. Do you want to disable the role and then delete it?`, "Yes", "No");
			if (disableRole === "Yes") {
				await this.changeState(item, false);
				await this.delete(item, true);
			}
			return;
		}

		if (hideConfirmation === false) {
			// If the user confirms the removal then delete the role.
			const deleteRole = await window.showWarningMessage(`Do you want to delete the App Role ${item.label}?`, "Yes", "No");
			if (deleteRole === "No") {
				return;
			}
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Deleting App Role...", item);

		// Get the parent application so we can read the app roles.
		const roles = await this.getAppRoles(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (roles === undefined) {
			return;
		}

		// Remove the app role from the array.
		roles.splice(
			roles.findIndex((r) => r.id === item.value!),
			1
		);

		// Update the application.
		await this.updateApplication(item.objectId!, { appRoles: roles }, status, "APP-ROLES");
	}

	// Captures the display name
	async inputDisplayName(title: string, existingValue: string | undefined, validation: (value: string) => string | undefined): Promise<string | undefined> {
		return await window.showInputBox({
			prompt: "App Role Display Name",
			placeHolder: "Enter a display name for the App Role",
			ignoreFocusOut: true,
			title: title,
			value: existingValue,
			validateInput: (value) => validation(value)
		});
	}

	// Captures the value
	async inputValue(title: string, isEditing: boolean, roles: AppRole[], existingValue: string | undefined, validation: (value: string, isEditing: boolean, oldValue: string | undefined, roles: AppRole[]) => string | undefined): Promise<string | undefined> {
		const debouncedValidation = debounce(validation, 500);
		// Prompt the user for the new value.
		return await window.showInputBox({
			prompt: "App Role Value",
			placeHolder: "Enter a value for the App Role",
			title: title,
			ignoreFocusOut: true,
			value: existingValue,
			validateInput: async (value) => debouncedValidation(value, isEditing, existingValue, roles)
		});
	}

	// Captures the description
	async inputDescription(title: string, existingValue: string | undefined, validation: (value: string) => string | undefined): Promise<string | undefined> {
		return await window.showInputBox({
			prompt: "App Role Description",
			placeHolder: "Enter a description for the App Role",
			title: title,
			ignoreFocusOut: true,
			value: existingValue,
			validateInput: (value) => validation(value)
		});
	}

	// Captures the allowed member types
	private async inputAllowedMemberTypes(title: string): Promise<{ label: string; description: string; value: string[] } | undefined> {
		return await window.showQuickPick(
			[
				{
					label: "Users/Groups",
					description: "Users and groups can be assigned this role.",
					value: ["User"]
				},
				{
					label: "Applications",
					description: "Only applications can be assigned this role.",
					value: ["Application"]
				},
				{
					label: "Both (Users/Groups + Applications)",
					description: "Users, groups and applications can be assigned this role.",
					value: ["User", "Application"]
				}
			],
			{
				placeHolder: "Select allowed member types",
				title: title,
				ignoreFocusOut: true
			}
		);
	}

	// Captures the details for an app role.
	private async inputRoleDetails(role: AppRole, isEditing: boolean, roles: AppRole[]): Promise<AppRole | undefined> {
		// Prompt the user for the new display name.
		const displayName = await this.inputDisplayName(isEditing === true ? "Edit App Role (1/5)" : "Add App Role (1/5)", role.displayName ?? undefined, validateAppRoleDisplayName);

		// If escape is pressed then return undefined.
		if (displayName === undefined) {
			return undefined;
		}

		const value = await this.inputValue(isEditing === true ? "Edit App Role (2/5)" : "Add App Role (2/5)", true, roles, role.value!, validateAppRoleValue);

		// If escape is pressed then return undefined.
		if (value === undefined) {
			return undefined;
		}

		// Prompt the user for the new display name.
		const description = await this.inputDescription(isEditing === true ? "Edit App Role (3/5)" : "Add App Role (3/5)", role.description ?? undefined, validateAppRoleDescription);

		// If escape is pressed then return undefined.
		if (description === undefined) {
			return undefined;
		}

		// Prompt the user for the new allowed member types.
		const allowed = await this.inputAllowedMemberTypes(isEditing === true ? "Edit App Role (4/5)" : "Add App Role (4/5)");

		// If escape is pressed or the new allowed member types is empty then return undefined.
		if (allowed === undefined) {
			return undefined;
		}

		// Prompt the user for the new state.
		const state = await window.showQuickPick(
			[
				{
					label: "Enabled",
					value: true
				},
				{
					label: "Disabled",
					value: false
				}
			],
			{
				placeHolder: "Select role state",
				title: isEditing === true ? "Edit App Role (5/5)" : "Add App Role (5/5)",
				ignoreFocusOut: true
			}
		);

		// If escape is pressed or the new state is empty then return undefined.
		if (state === undefined) {
			return undefined;
		}

		// Set the new values on the app role.
		role.id = role.id ?? uuidv4();
		role.displayName = displayName;
		role.value = value;
		role.description = description;
		role.allowedMemberTypes = allowed.value;
		role.isEnabled = state.value;

		return role;
	}

	// Gets the app roles for an application registration.
	private async getAppRoles(id: string): Promise<AppRole[] | undefined> {
		const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(id, "appRoles");
		if (result.success === true && result.value !== undefined) {
			return result.value?.appRoles;
		} else {
			await this.handleError(result.error);
			return undefined;
		}
	}
}
