# All Documents Web Part

## Summary

The **All Documents Web Part** is a SharePoint Framework (SPFx) solution that aggregates and displays documents from all document libraries within a SharePoint site. It provides advanced filtering, search capabilities, and customizable column configurations to help users efficiently manage and navigate their document collections.

![SPFx Version](https://img.shields.io/badge/SPFx-1.21.1-green.svg)
![Node.js Version](https://img.shields.io/badge/Node.js-22.x-green.svg)

## Key Features

- **📚 Multi-Library Aggregation**: Automatically loads documents from all document libraries in the current site
- **🔍 Advanced Search**: Real-time search functionality across all document names
- **🎯 Smart Filtering**: Filter documents by library and custom metadata columns
- **📊 Customizable Columns**: Configure which system and custom columns to display
- **🎨 Fluent UI Design**: Seamless integration with SharePoint's modern UI
- **🔗 Direct Document Access**: Open documents directly in edit mode with proper URL handling
- **📁 File Type Icons**: Visual file type indicators for easy document identification
- **🔄 Sortable Columns**: Click any column header to sort documents
- **🔒 Site-Restricted**: Security control to limit deployment to authorized sites only

## Screenshots

![All Documents Web Part in action](./assets/screenshot.png)

*Documents displayed with filtering and search capabilities*

## Prerequisites

Before deploying this solution, ensure you have:

- **SharePoint Online** tenant with appropriate permissions
- **Node.js** version 22.14.0 or higher (< 23.0.0)
- **SharePoint Framework** development environment set up
- Admin access to the SharePoint App Catalog

## Installation & Deployment

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/all-documents-web-part.git
cd all-documents-web-part
```

### Step 2: Install Dependencies

```bash
npm install
```

### Step 3: Build the Solution

```bash
gulp clean
gulp build
gulp bundle --ship
gulp package-solution --ship
```

### Step 4: Deploy to SharePoint

1. Navigate to your SharePoint App Catalog site
2. Upload the `.sppkg` file from the `sharepoint/solution` folder
3. When prompted, click **Deploy**
4. Trust the solution when asked

### Step 5: Add to a SharePoint Page

1. Navigate to the target SharePoint site (must be in the authorized list)
2. Edit a page and click **Add a web part**
3. Search for **AllDocuments**
4. Add the web part to your page
5. Configure the web part properties

## Configuration

### Web Part Properties

Access the configuration panel by clicking the **Edit** (pencil) icon on the web part.

#### System Columns

Toggle visibility for built-in columns:

- **Show Modified**: Display the last modified date
- **Show Modified By**: Display who last modified the document
- **Show Library**: Display which library contains the document

#### Custom Columns

Define custom metadata columns to display:

**Format**: `InternalName,Label;InternalName,Label;...`

**Example**:
```
Categoria,Category;Estado,Status;Departamento,Department
```

- `InternalName`: The SharePoint internal column name (case-sensitive)
- `Label`: The display name shown in the web part header

**Tips**:
- Use semicolons (`;`) to separate multiple columns
- Use commas (`,`) to separate internal name from display label
- If no label is provided, the internal name will be used as the label

### Authorized Sites

The web part includes a security feature that restricts deployment to specific sites. By default, it's configured for:

```typescript
const allowedSites = ["sites/sp-FIN"];
```

To modify authorized sites:

1. Open `src/webparts/allDocuments/components/AllDocuments.tsx`
2. Update the `allowedSites` array in the `render()` method
3. Rebuild and redeploy the solution

## Usage Guide

### Searching for Documents

Use the search box at the top to filter documents by name in real-time. The search is case-insensitive and filters across all loaded documents.

### Filtering Documents

If custom columns or the Library column are enabled, dropdown filters will appear below the search box. Select values from these dropdowns to narrow down the document list.

### Sorting Documents

Click any column header to sort documents by that column. Click again to reverse the sort order. A visual indicator (▲ or ▼) shows the current sort column and direction.

### Opening Documents

Click any document name to open it:
- **Office documents** (Word, Excel, PowerPoint): Opens in Office Online editor
- **PDF files**: Opens in SharePoint document library viewer
- **Other files**: Downloads or displays based on file type

## Technical Architecture

### Technology Stack

- **Framework**: SharePoint Framework (SPFx) 1.21.1
- **UI Library**: Fluent UI React 8.123.0
- **Icons**: @fluentui/react-file-type-icons
- **Data Access**: SharePoint REST API via SPHttpClient
- **Build Tools**: Gulp, TypeScript 5.3.3

### Component Structure

```
src/
├── webparts/
│   └── allDocuments/
│       ├── AllDocumentsWebPart.ts          # Web part entry point
│       ├── components/
│       │   ├── AllDocuments.tsx            # Main React component
│       │   ├── AllDocuments.module.scss    # Styles
│       │   └── IAllDocumentsProps.ts       # Props interface
│       ├── loc/                            # Localization resources
│       └── pnpjsConfig.ts                  # PnP JS configuration
```

### Data Flow

1. **Component Mount**: On load, the web part queries all document libraries in the site
2. **Data Retrieval**: For each library, it fetches all documents with metadata using CAML queries
3. **Data Processing**: Documents are aggregated, deduplicated, and enriched with custom column data
4. **State Management**: React state manages the document list, filters, and UI state
5. **Rendering**: DetailsList component displays the data with sorting and filtering applied

### Key REST API Calls

**Get All Document Libraries:**
```
GET /_api/web/lists?$filter=BaseTemplate eq 101&$select=Id,Title
```

**Get Documents from a Library:**
```
POST /_api/web/lists(guid'<LibraryId>')/RenderListDataAsStream
```

## Development

### Local Development Server

Run the web part in your local development environment:

```bash
gulp serve
```

This will open the SharePoint Workbench where you can test the web part.

### Debug Configuration

The solution is configured to serve on **port 4321** with HTTPS enabled. Update `config/serve.json` to change these settings.

### Code Quality

The project includes ESLint configuration for code quality:

```bash
npm run build  # Runs ESLint checks
```

## Troubleshooting

### Common Issues

**Issue**: Web part shows "not authorized" message
- **Solution**: Verify the site URL is in the `allowedSites` array in `AllDocuments.tsx`

**Issue**: Custom columns not appearing
- **Solution**: 
  - Ensure internal names are correct (case-sensitive)
  - Verify the columns exist in all document libraries
  - Check browser console for errors

**Issue**: File URLs are malformed
- **Solution**: This has been fixed in version 1.0.2.50 - ensure you're using the latest version

**Issue**: File type icons not showing correctly
- **Solution**: Icons are now based on file extension, not file name - update to the latest version

**Issue**: Documents not loading
- **Solution**: 
  - Check browser console for API errors
  - Verify user has read permissions on all libraries
  - Ensure the site collection has document libraries

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.2.70 | 2025-11-11 | - Added customization option to hide system columns: Modified, modifiedby and library |
| 1.0.2.60 | 2025-10-11 | - Added way to view customcolumns sharepoint formatting with colors from Sharpeoint |
| 1.0.2.50 | 2025-10-11 | - Added system column toggles (Modified, Modified By, Library)<br>- Fixed file URL generation issue<br>- Improved file type icon consistency<br>- Added library filter dropdown |
| 1.0.2.00 | 2025-10-11 | - Changed UI from custom CSS to FluidUI 8 for improved consistency with Sharepoint <br> - Added File Icons|
| 1.0.1.97 | 2025-30-06 | - Added loading until fully loaded <br> - Added block to only load on specific site|
| 1.0.1.90 | 2025-30-06 | Fixed the ability to open pdf files |
| 1.0.1.80 | 2025-10-06 | Added ability to sort items |
| 1.0.1.50 | 2025-10-06 | Added the ability to add CustomColumns via WebPart config in the site |
| 1.0.1.00 | 2025-10-06 | Added custom css styling |
| 1.0.0.80 | 2025-10-06 | Added ability to add CustomColumns via hardcode |
| 1.0.0.00 | 2025-10-06 | Initial release |

## Browser Support

- ✅ Microsoft Edge (Chromium)
- ✅ Google Chrome
- ✅ Mozilla Firefox
- ✅ Safari (macOS)

## Performance Considerations

- The web part loads all documents from all libraries on mount
- For sites with thousands of documents, initial load may take several seconds
- Consider implementing pagination for very large document collections
- Filtering and sorting are performed client-side for instant results

## Security & Permissions

- Users can only see documents they have permission to access
- The web part respects SharePoint item-level security
- Site restriction feature prevents unauthorized deployment
- No elevated permissions are required

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Author

**Oleksandr Lugovskyy Rodriguez**
- LinkedIn: [Oleksandr Lugovskyy Rodriguez](https://www.linkedin.com/in/oleksandr-lugovskyy-rodriguez/)

## Acknowledgments

- Built with [SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- UI components from [Fluent UI](https://developer.microsoft.com/en-us/fluentui)
- Icons from [@fluentui/react-file-type-icons](https://www.npmjs.com/package/@fluentui/react-file-type-icons)

## Resources

- [SharePoint Framework Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Fluent UI React Components](https://developer.microsoft.com/en-us/fluentui#/controls/web)
- [SharePoint REST API Reference](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)

## Support

For issues, questions, or feature requests:
1. Check the [Troubleshooting](#troubleshooting) section
2. Review existing [GitHub Issues](https://github.com/gyswu/all-documents-web-part/issues)
3. Create a new issue with detailed information

---

Made with assistance from Chat-GPT and Claude (Shotout to claude 4.5 that does not mess the code)