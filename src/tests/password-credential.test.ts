import * as vscode from "vscode";
import * as validation from "../utils/validation";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { PasswordCredentialService } from "../services/password-credential";
import { mockApplications, mockAppObjectId, mockNewPasswordKeyId, seedMockData } from "./data/test-data";
import { getTopLevelTreeItem } from "./test-utils";
import { format } from "date-fns";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for password credential service
describe("Password Credential Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const passwordCredentialService = new PasswordCredentialService(graphApiRepository, treeDataProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerTreeErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let iconSpy: jest.SpyInstance<any, [id: string, color?: any | undefined], any>;

	// Create variables used in the tests
	let item: AppRegItem;

	beforeAll(async () => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		// Reset mock data
		seedMockData();

		//Restore the default mock implementations
		jest.restoreAllMocks();

		// Define a standard mock implementation for the dialog functions
		vscode.window.showWarningMessage = jest.fn().mockResolvedValue("Yes");
		vscode.window.showInputBox = jest.fn().mockResolvedValue("Test Input");

		// Define spies on the functions to be tested
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(passwordCredentialService), "triggerRefresh");
		triggerTreeErrorSpy = jest.spyOn(Object.getPrototypeOf(treeDataProvider), "handleError");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(passwordCredentialService), "handleError");

		// The item to be tested
		item = { objectId: mockAppObjectId, contextValue: "PASSWORD-CREDENTIALS" };
	});

	afterAll(() => {
		// Dispose of the application service
		passwordCredentialService.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(passwordCredentialService).toBeDefined();
	});

	test("Delete password credential successfully", async () => {
		// Act
		await passwordCredentialService.delete(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "PASSWORD-CREDENTIALS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(0);
	});

	test("Delete password credential with error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "deletePasswordCredential").mockImplementation(async () => ({ success: false, error: new Error("Test Error") }));

		// Act
		await passwordCredentialService.delete(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Add password credential successfully", async () => {
		// Arrange
		jest.spyOn(passwordCredentialService, "inputDescription").mockImplementation(async () => "Test Description");
		jest.spyOn(passwordCredentialService, "inputExpiryDate").mockImplementation(async (expiryDate: Date, validation: (value: string) => string | undefined) => {
			const expiry = format(expiryDate, "yyyy-MM-dd");
			return validation(expiry) === undefined ? expiry : undefined;
		});

		// Act
		await passwordCredentialService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "PASSWORD-CREDENTIALS");
		const passwordItem = treeItem?.children?.find((item) => item.value === mockNewPasswordKeyId);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(vscode.env.clipboard.readText()).toEqual("NEWPASSWORD");
		expect(passwordItem).toBeDefined();
		expect(passwordItem?.children?.length).toEqual(3);
		expect(treeItem?.children?.length).toEqual(2);
	});

	test("Add password credential with incorrect date validation bad characters error", async () => {
		// Arrange
		const validationSpy = jest.spyOn(validation, "validatePasswordCredentialExpiryDate");
		jest.spyOn(passwordCredentialService, "inputDescription").mockImplementation(async () => "Test Description");
		jest.spyOn(passwordCredentialService, "inputExpiryDate").mockImplementation(async (_expiryDate: Date, validation: (value: string) => string | undefined) => {
			const result = validation("NOTADATE");
			expect(result).toBe("Expiry must be in the format YYYY-MM-DD or YYYY/MM/DD.");
			return result;
		});

		// Act
		await passwordCredentialService.add(item);

		// Assert
		expect(passwordCredentialService.inputExpiryDate).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Add password credential with incorrect date validation not a date error", async () => {
		// Arrange
		const validationSpy = jest.spyOn(validation, "validatePasswordCredentialExpiryDate");
		jest.spyOn(passwordCredentialService, "inputDescription").mockImplementation(async () => "Test Description");
		jest.spyOn(passwordCredentialService, "inputExpiryDate").mockImplementation(async (_expiryDate: Date, validation: (value: string) => string | undefined) => {
			const result = validation("2100-01-42");
			expect(result).toBe("Expiry must be a valid date.");
			return result;
		});

		// Act
		await passwordCredentialService.add(item);

		// Assert
		expect(passwordCredentialService.inputExpiryDate).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Add password credential with past date validation error", async () => {
		// Arrange
		const validationSpy = jest.spyOn(validation, "validatePasswordCredentialExpiryDate");
		jest.spyOn(passwordCredentialService, "inputDescription").mockImplementation(async () => "Test Description");
		jest.spyOn(passwordCredentialService, "inputExpiryDate").mockImplementation(async (_expiryDate: Date, validation: (value: string) => string | undefined) => {
			const result = validation("2000-01-01");
			expect(result).toBe("Expiry must be in the future.");
			return result;
		});

		// Act
		await passwordCredentialService.add(item);

		// Assert
		expect(passwordCredentialService.inputExpiryDate).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Add password credential with future date validation error", async () => {
		// Arrange
		const validationSpy = jest.spyOn(validation, "validatePasswordCredentialExpiryDate");
		jest.spyOn(passwordCredentialService, "inputDescription").mockImplementation(async () => "Test Description");
		jest.spyOn(passwordCredentialService, "inputExpiryDate").mockImplementation(async (_expiryDate: Date, validation: (value: string) => string | undefined) => {
			const result = validation("3000-01-01");
			expect(result).toBe("Expiry must be less than 2 years in the future.");
			return result;
		});

		// Act
		await passwordCredentialService.add(item);

		// Assert
		expect(passwordCredentialService.inputExpiryDate).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Add password credential with error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "addPasswordCredential").mockImplementation(async () => ({ success: false, error: new Error("Test Error") }));

		// Act
		await passwordCredentialService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Error getting password credential children", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "PASSWORD-CREDENTIALS" };
		const error = new Error("Error getting password credential children");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (id: string, select: string) => {
			if (select === "passwordCredentials") {
				return { success: false, error };
			}
			return mockApplications.find((app) => app.id === id);
		});

		// Act
		await treeDataProvider.render();
		await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "PASSWORD-CREDENTIALS");

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});
});
