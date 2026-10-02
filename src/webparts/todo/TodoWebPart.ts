import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "TodoWebPartStrings";
import Todo from "./components/Todo";
import { ITodoProps } from "./components/ITodoProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ITodoItem } from "./models/ITodoItem";

export interface ITodoItemsResponse {
  value: ITodoItem[];
}

export interface ITodoWebPartProps {
  description: string;
}

export default class TodoWebPart extends BaseClientSideWebPart<ITodoWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  private _getTodoItems(): Promise<ITodoItemsResponse> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/GetByTitle('TodoTasks')/items` +
      `?$select=Id,Title,Completed`;

    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(
            `Failed to load todo items. Status: ${response.status}`,
          );
        }

        return response.json();
      });
  }

  private _addTodoItem(title: string): Promise<void> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/GetByTitle('TodoTasks')/items`;

    const body = {
      Title: title,
      Completed: false,
    };

    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
        },

        body: JSON.stringify(body),
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(
            `Failed to add todo item. Status: ${response.status}`,
          );
        }
      });
  }
  private _updateTodoCompleted(id: number, completed: boolean): Promise<void> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/GetByTitle('TodoTasks')/items(${id})`;

    const body = {
      Completed: completed,
    };

    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",

          // MERGE means update only the fields we send.
          "X-HTTP-Method": "MERGE",

          // "*" means update regardless of the current item ETag/version.
          "IF-MATCH": "*",
        },

        body: JSON.stringify(body),
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(
            `Failed to update todo item. Status: ${response.status}`,
          );
        }
      });
  }

  private _deleteTodoItem(id: number): Promise<void> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/GetByTitle('TodoTasks')/items(${id})`;

    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",

          // Tell SharePoint this POST should behave as DELETE.
          "X-HTTP-Method": "DELETE",

          // "*" means ignore the current ETag/version.
          "IF-MATCH": "*",
        },
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(
            `Failed to delete todo item. Status: ${response.status}`,
          );
        }
      });
  }

  private _updateTodoTitle(id: number, title: string): Promise<void> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists/GetByTitle('TodoTasks')/items(${id})`;

    const body = {
      Title: title,
    };

    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
        },

        body: JSON.stringify(body),
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(
            `Failed to update todo title. Status: ${response.status}`,
          );
        }
      });
  }

  public render(): void {
    // First get Todo items from the SharePoint list.
    this._getTodoItems()
      .then((response: ITodoItemsResponse) => {
        // response.value contains the actual SharePoint list items.
        const todoItems: ITodoItem[] = response.value;

        // Create our React Todo component.
        const element: React.ReactElement<ITodoProps> = React.createElement(
          Todo,
          {
            description: this.properties.description,
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            userDisplayName: this.context.pageContext.user.displayName,

            // Pass SharePoint items from SPFx into React.
            todoItems: todoItems,
            onAddTodo: async (title: string): Promise<void> => {
              await this._addTodoItem(title);
              this.render();
            },
            onToggleTodo: async (
              id: number,
              completed: boolean,
            ): Promise<void> => {
              await this._updateTodoCompleted(id, completed);

              // Reload latest list data after update.
              this.render();
            },
            onDeleteTodo: async (id: number): Promise<void> => {
              await this._deleteTodoItem(id);

              // Reload the latest SharePoint items after delete.
              this.render();
            },
            onEditTodo: async (id: number, title: string): Promise<void> => {
              await this._updateTodoTitle(id, title);

              this.render();
            },
          },
        );

        // Render the React component inside the SPFx web-part container.
        ReactDom.render(element, this.domElement);
      })
      .catch((error: Error) => {
        // For now log API errors so we can inspect them.
        console.error("Error loading TodoTasks:", error);
      });
  }
  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment,
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null,
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null,
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
