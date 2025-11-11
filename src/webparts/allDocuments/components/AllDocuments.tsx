import * as React from "react";
import { IAllDocumentsProps } from "./IAllDocumentsProps";
import { SPHttpClient } from "@microsoft/sp-http";
import styles from "./AllDocuments.module.scss";
import {
  initializeFileTypeIcons,
  getFileTypeIconProps,
} from "@fluentui/react-file-type-icons";
import { Icon } from "@fluentui/react/lib/Icon";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from "@fluentui/react/lib/DetailsList";
import { SearchBox } from "@fluentui/react/lib/SearchBox";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { Stack } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Link } from "@fluentui/react/lib/Link";
import { mergeStyles } from "@fluentui/react/lib/Styling";

// Initialize the icons once
initializeFileTypeIcons();

export interface IDocumentItem {
  name: string;
  extension: string;
  modified: string;
  modifiedBy: string;
  library: string;
  editUrl: string;
  customColumns: { [key: string]: string };
  columnFormatting: { [key: string]: string }; // HTML string from SharePoint
}

interface IState {
  items: IDocumentItem[];
  loading: boolean;
  filters: { [key: string]: string };
  filterOptions: { [key: string]: Set<string> };
  searchTerm: string;
  columns: IColumn[];
  columnFormattersByLibrary: {
    [libId: string]: { [fieldName: string]: Record<string, unknown> };
  };
}

const iconClass = mergeStyles({
  marginRight: 8,
  verticalAlign: "middle",
  fontSize: 24,
  height: 24,
  width: 24,
  display: "inline-block",
});

export default class AllDocuments extends React.Component<IAllDocumentsProps, IState> {
  constructor(props: IAllDocumentsProps) {
    super(props);
    this.state = {
      items: [],
      loading: true,
      filters: {},
      filterOptions: {},
      searchTerm: "",
      columns: this._buildColumns(),
      columnFormattersByLibrary: {},
    };
  }

  public async componentDidMount(): Promise<void> {
    try {
      const res = await this.props.spHttpClient.get(
        `${this.props.siteUrl}/_api/web/lists?$filter=BaseTemplate eq 101&$select=Id,Title`,
        SPHttpClient.configurations.v1
      );
      const json = await res.json();
      const libraries = json.value;

      const allItems: IDocumentItem[] = [];
      const filterOptions: { [key: string]: Set<string> } = {};
      const columnFormattersByLibrary: {
        [libId: string]: { [fieldName: string]: Record<string, unknown> };
      } = {};

      // Load column formatting if enabled
      if (this.props.useColumnFormatting) {
        for (const lib of libraries) {
          columnFormattersByLibrary[lib.Id] = {};

          for (const col of this.props.customColumns) {
            try {
              const fieldRes = await this.props.spHttpClient.get(
                `${this.props.siteUrl}/_api/web/lists(guid'${lib.Id}')/fields?$filter=InternalName eq '${col.internalName}'&$select=CustomFormatter`,
                SPHttpClient.configurations.v1
              );
              const fieldJson = await fieldRes.json();

              if (
                fieldJson.value &&
                fieldJson.value.length > 0 &&
                fieldJson.value[0].CustomFormatter
              ) {
                try {
                  const formatter = JSON.parse(
                    fieldJson.value[0].CustomFormatter
                  );
                  columnFormattersByLibrary[lib.Id][col.internalName] =
                    formatter;
                  console.log(
                    `Loaded formatter for ${col.internalName}:`,
                    formatter
                  );
                } catch (parseErr) {
                  console.warn(
                    `Could not parse formatter for ${col.internalName}:`,
                    parseErr
                  );
                }
              }
            } catch (err) {
              console.warn(
                `Could not load formatting for column ${col.internalName}:`,
                err
              );
            }
          }
        }
      }

      for (const lib of libraries) {
        const camlQuery = {
          ViewXml: `
            <View Scope='RecursiveAll'>
              <Query></Query>
              <ViewFields>
                <FieldRef Name='FileLeafRef' />
                <FieldRef Name='FileRef' />
                <FieldRef Name='Modified' />
                <FieldRef Name='Editor' />
                ${this.props.customColumns
                  .map((col) => `<FieldRef Name='${col.internalName}' />`)
                  .join("")}
              </ViewFields>
            </View>`,
        };

        const itemsRes = await this.props.spHttpClient.post(
          `${this.props.siteUrl}/_api/web/lists(guid'${lib.Id}')/RenderListDataAsStream`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=verbose",
            },
            body: JSON.stringify({ parameters: camlQuery }),
          }
        );

        const itemsJson = await itemsRes.json();
        const rows = itemsJson?.Row;
        if (!rows || rows.length === 0) {
          console.warn(`No items found in library: ${lib.Title}`);
          continue;
        }

        for (const file of rows) {
          if (file.FSObjType === "1" || file.FSObjType === 1) continue;

          const fileName = file.FileLeafRef;
          const filePath = file.FileRef;
          const modified = file.Modified;
          const editor =
            file.Editor?.[0]?.title || file.Editor?.title || file.Editor || "";
          const extension =
            fileName.split(".").pop()?.toLowerCase() || "unknown";

          const customData: { [key: string]: string } = {};
          const columnFormatting: { [key: string]: string } = {};

          for (const col of this.props.customColumns) {
            const value = file[col.internalName];
            customData[col.internalName] = value || "";
            if (!filterOptions[col.internalName])
              filterOptions[col.internalName] = new Set<string>();
            if (value) filterOptions[col.internalName].add(value);

            // Apply column formatting if enabled
            if (
              this.props.useColumnFormatting &&
              columnFormattersByLibrary[lib.Id]?.[col.internalName]
            ) {
              const htmlString = this._renderColumnFormatting(
                columnFormattersByLibrary[lib.Id][col.internalName],
                value,
                col.internalName
              );
              columnFormatting[col.internalName] = htmlString;
            }
          }

          const officeExtensions = [
            "docx",
            "xlsx",
            "pptx",
            "doc",
            "xls",
            "ppt",
          ];
          let editUrl: string;

          if (officeExtensions.includes(extension)) {
            editUrl = `${
              this.props.siteUrl
            }/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(
              filePath
            )}&action=edit&mobileredirect=true`;
          } else if (extension === "pdf") {
            const parentPath = filePath.substring(0, filePath.lastIndexOf("/"));
            const libPath = `${parentPath}/Forms/AllItems.aspx`;
            editUrl = `${
              window.location.origin
            }${libPath}?id=${encodeURIComponent(
              filePath
            )}&parent=${encodeURIComponent(parentPath)}`;
          } else {
            // Fix for file URLs - ensure no double site paths
            editUrl = `${window.location.origin}${filePath}`;
          }

          allItems.push({
            name: fileName,
            extension: extension,
            modified: modified,
            modifiedBy: editor,
            library: lib.Title,
            editUrl: editUrl,
            customColumns: customData,
            columnFormatting: columnFormatting,
          });
        }
      }

      this.setState({
        items: allItems,
        loading: false,
        filters: {},
        filterOptions: filterOptions,
        columnFormattersByLibrary: columnFormattersByLibrary,
      });
      //DEBUG PURPOSE
      //console.log("Loaded document items:", allItems);
    } catch (err) {
      console.error("Error loading documents:", err);
      this.setState({ loading: false });
    }
  }

  private _renderColumnFormatting(
    formatter: Record<string, unknown>,
    currentValue: string,
    fieldName: string
  ): string {
    try {
      return this._processElement(formatter, currentValue, fieldName);
    } catch (err) {
      console.warn("Error rendering column formatting:", err);
      return currentValue;
    }
  }

  private _processElement(
    element: Record<string, unknown>,
    currentValue: string,
    fieldName: string
  ): string {
    if (!element || typeof element !== "object") {
      return "";
    }

    const elmType = element.elmType as string;
    const children = element.children as Record<string, unknown>[] | undefined;
    const txtContent = element.txtContent as string | undefined;
    const attributes = element.attributes as
      | Record<string, unknown>
      | undefined;
    const style = element.style as Record<string, string> | undefined;

    // Build the HTML element
    let html = `<${elmType || "div"}`;

    // Process attributes (including class)
    if (attributes) {
      if (attributes.class) {
        const classValue = this._evaluateOperator(
          attributes.class,
          currentValue,
          fieldName
        );
        if (classValue) {
          html += ` class="${classValue}"`;
        }
      }
    }

    // Process inline styles
    if (style) {
      const styleStr = Object.entries(style)
        .map(([key, value]) => {
          const cssKey = key.replace(/([A-Z])/g, "-$1").toLowerCase();
          return `${cssKey}: ${value}`;
        })
        .join("; ");
      if (styleStr) {
        html += ` style="${styleStr}"`;
      }
    }

    html += ">";

    // Process text content
    if (txtContent) {
      const text = txtContent.replace(/\[\$(\w+)\]/g, (match, field) => {
        return field === fieldName ? currentValue : match;
      });
      html += text;
    }

    // Process children recursively
    if (children && Array.isArray(children)) {
      for (const child of children) {
        html += this._processElement(child, currentValue, fieldName);
      }
    }

    html += `</${elmType || "div"}>`;

    return html;
  }

  private _evaluateOperator(
    operatorObj: unknown,
    currentValue: string,
    fieldName: string
  ): string {
    if (typeof operatorObj === "string") {
      return operatorObj;
    }

    if (typeof operatorObj !== "object" || operatorObj === null) {
      return "";
    }

    const obj = operatorObj as Record<string, unknown>;
    const operator = obj.operator as string;
    const operands = obj.operands as unknown[];

    if (!operator || !operands) {
      return "";
    }

    if (operator === ":") {
      // Ternary operator: condition ? trueValue : falseValue
      if (operands.length >= 3) {
        const condition = this._evaluateOperator(
          operands[0],
          currentValue,
          fieldName
        );
        if (condition === "true") {
          return this._evaluateOperator(operands[1], currentValue, fieldName);
        } else {
          return this._evaluateOperator(operands[2], currentValue, fieldName);
        }
      }
    } else if (operator === "==") {
      // Equality check
      if (operands.length >= 2) {
        const left = this._evaluateOperand(
          operands[0],
          currentValue,
          fieldName
        );
        const right = this._evaluateOperand(
          operands[1],
          currentValue,
          fieldName
        );
        return left === right ? "true" : "false";
      }
    }

    return "";
  }

  private _evaluateOperand(
    operand: unknown,
    currentValue: string,
    fieldName: string
  ): string {
    if (typeof operand === "string") {
      // Handle field references like [$FieldName]
      if (operand.startsWith("[$") && operand.endsWith("]")) {
        const field = operand.substring(2, operand.length - 1);
        return field === fieldName ? currentValue : "";
      }
      return operand;
    }

    if (typeof operand === "object" && operand !== null) {
      return this._evaluateOperator(operand, currentValue, fieldName);
    }

    return String(operand || "");
  }

  private _buildColumns(): IColumn[] {
    const columns: IColumn[] = [
      {
        key: "name",
        name: "Name",
        fieldName: "name",
        minWidth: 250,
        maxWidth: 400,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocumentItem) => {
          return (
            <Stack horizontal verticalAlign="center">
              <Icon
                {...getFileTypeIconProps({
                  extension: item.extension,
                  size: 24,
                  imageFileType: "svg",
                })}
                className={iconClass}
              />
              <Link
                href={item.editUrl}
                target="_blank"
                styles={{
                  root: {
                    whiteSpace: "nowrap",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                  },
                }}
              >
                {item.name}
              </Link>
            </Stack>
          );
        },
      },
    ];

    // Add system columns based on configuration
    if (this.props.showModified) {
      columns.push({
        key: "modified",
        name: "Modified",
        fieldName: "modified",
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocumentItem) => {
          const date = new Date(item.modified);
          return <Text>{date.toLocaleString()}</Text>;
        },
      });
    }

    if (this.props.showModifiedBy) {
      columns.push({
        key: "modifiedBy",
        name: "Modified By",
        fieldName: "modifiedBy",
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocumentItem) => {
          return <Text>{item.modifiedBy}</Text>;
        },
      });
    }

    if (this.props.showLibrary) {
      columns.push({
        key: "library",
        name: "Library",
        fieldName: "library",
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocumentItem) => {
          return <Text>{item.library}</Text>;
        },
      });
    }

    // Add custom columns
    this.props.customColumns.forEach((col) => {
      columns.push({
        key: col.internalName,
        name: col.label,
        fieldName: col.internalName,
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocumentItem) => {
          const value = item.customColumns[col.internalName] || "";

          if (
            this.props.useColumnFormatting &&
            item.columnFormatting[col.internalName]
          ) {
            // Render HTML with SharePoint classes
            return (
              <div
                dangerouslySetInnerHTML={{
                  __html: item.columnFormatting[col.internalName],
                }}
              />
            );
          }

          return <Text>{value}</Text>;
        },
      });
    });

    return columns;
  }

  private _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];

    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });

    const newItems = this._copyAndSort(
      items,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );

    this.setState({
      columns: newColumns,
      items: newItems,
    });
  };

  private _copyAndSort(
    items: IDocumentItem[],
    columnKey: string,
    isSortedDescending?: boolean
  ): IDocumentItem[] {
    return items.slice(0).sort((a: IDocumentItem, b: IDocumentItem) => {
      let aValue: string;
      let bValue: string;

      if (
        columnKey === "name" ||
        columnKey === "extension" ||
        columnKey === "modified" ||
        columnKey === "modifiedBy" ||
        columnKey === "library"
      ) {
        aValue = a[columnKey as keyof IDocumentItem] as string;
        bValue = b[columnKey as keyof IDocumentItem] as string;
      } else {
        // Custom column
        aValue = a.customColumns[columnKey] || "";
        bValue = b.customColumns[columnKey] || "";
      }

      if (isSortedDescending) {
        return aValue.localeCompare(bValue) * -1;
      } else {
        return aValue.localeCompare(bValue);
      }
    });
  }

  private _onSearchChange = (
    event?: React.ChangeEvent<HTMLInputElement>,
    newValue?: string
  ): void => {
    this.setState({ searchTerm: newValue || "" });
  };

  private _onFilterChanged = (
    column: string,
    option?: IDropdownOption
  ): void => {
    const newFilters = {
      ...this.state.filters,
      [column]: (option?.key as string) || "",
    };
    this.setState({ filters: newFilters });
  };

  private _getFilteredItems(): IDocumentItem[] {
    const { items, filters, searchTerm } = this.state;

    let filtered = items.filter((item) =>
      Object.entries(filters).every(
        ([key, val]) => val === "" || item.customColumns[key] === val
      )
    );

    if (searchTerm) {
      filtered = filtered.filter((item) =>
        item.name.toLowerCase().includes(searchTerm.toLowerCase())
      );
    }

    return filtered;
  }

  public render(): React.ReactElement<IAllDocumentsProps> {
    const { loading, filterOptions, filters, searchTerm } = this.state;
    const allowedSites = ["sites/sp-FIN"];

    const authorized = allowedSites.some((path) =>
      this.props.siteUrl.includes(path)
    );

    if (!authorized) {
      return (
        <MessageBar messageBarType={MessageBarType.error}>
          ⚠️ This webpart is not authorized to be loaded in this site.
        </MessageBar>
      );
    }

    if (loading || this.props.customColumns.length === 0) {
      return (
        <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
          <Spinner size={SpinnerSize.large} label="Loading Files..." />
        </Stack>
      );
    }

    const filteredItems = this._getFilteredItems();

    return (
      <Stack tokens={{ childrenGap: 16 }} className={styles.container}>
        {/* Search Box */}
        <SearchBox
          placeholder="Search files..."
          value={searchTerm}
          onChange={this._onSearchChange}
          underlined
        />

        {/* Filters */}
        <Stack
          horizontal
          wrap
          tokens={{ childrenGap: 16 }}
          styles={{ root: { marginBottom: 8 } }}
        >
          {this.props.customColumns.map((col) => {
            const options: IDropdownOption[] = [
              { key: "", text: "All" },
              ...[...(filterOptions[col.internalName] || [])].map((val) => ({
                key: val,
                text: val,
              })),
            ];

            return (
              <Stack.Item
                key={col.internalName}
                styles={{ root: { minWidth: 200 } }}
              >
                <Dropdown
                  label={col.label}
                  selectedKey={filters[col.internalName] || ""}
                  onChange={(e, option) =>
                    this._onFilterChanged(col.internalName, option)
                  }
                  options={options}
                  styles={{ dropdown: { width: 200 } }}
                />
              </Stack.Item>
            );
          })}
        </Stack>

        {/* DetailsList */}
        <DetailsList
          items={filteredItems}
          columns={this.state.columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          isHeaderVisible={true}
          compact={false}
        />
      </Stack>
    );
  }
}
