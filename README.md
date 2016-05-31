# gulp-spsync
Gulp plugin for synchronizing local files with a SharePoint library

# Features
 
* Gulp plugin
* Copies local files to a SharePoint Document libraries and galleries

# How to use

1. Prepare SharePoint by registering a SharePoint app using appregnew.aspx. Eg go to https://contoso.sharepoint.com/sites/site/_layouts/15/appregnew.aspx
2. Click on Generate for both Client Id and Client Secret. For Title, App Domain and Redirect URI, write something you don't care about. Then click on Create
3. Note down the Client Id and Client Secret, you will need it later
4. Navigate to appinv.aspx, https://contoso.sharepoint.com/sites/site/_layouts/15/appinv.aspx, enter the client ID in the App Id box and press Lookup
5. In the Permission Request XML text box enter the following XML and click Create (Note: `FullControl` is required to update assets in the Master Page gallery)  
```xml
<AppPermissionRequests AllowAppOnlyPolicy="true">
    <AppPermissionRequest
        Scope="http://sharepoint/content/sitecollection/web"
        Right="FullControl"/>
</AppPermissionRequests>
```
6. In the following consent screen choose to trust the App by clicking on Trust It!
7. Open a folder using Visual studio code
8. Run `npm install gulp` to install the Gulp task runner
9. Run `npm install wictorwilen/gulp-spsync` to install to install the gulp-spsync 
10. Press Ctrl-Shift-P, type Task and choose to Configure Task Runner
11. In the tasks.json file that is being created replace the contents with the following:
```json
{
    "version": "0.1.0",
    "command": "gulp",
    "isShellCommand": true,
    "tasks": [
        {
            "taskName": "default",
            "isBuildCommand": true,
            "showOutput": "silent"
        }
    ]
}	
```
12. Create a new file in the root of your folder called `gulpfile.js`, and modify it as follows. This task will monitor all files in the `Src` folder
```javascript
var gulp = require('gulp')
var sp = require('gulp-spsync')
gulp.task('default', function() {
return gulp.src('src/**/*.*').
    pipe(sp({
        "client_id":"3d271647-2e12-4ae5-9271-04b3aa67dcd3",
        "client_secret":"Zk9ORywN0gaGljrtlxfp+s5vh7ZyWV4dRpOXCLjtl8U=",
        "realm" : "",
        "site" : "https://contoso.sharepoint.com/sites/site",
        "verbose": "true"
    })).		
    pipe(gulp.dest('build'))
})
```
13. Replace the client_id and client_secret parameters with the value for the App you just created
14. Replace the site URL with your site URL
15. Create a folder called `Src` (you can call it whatever you want, but the tasks above/below uses `Src`)
16. Create sub folders to the `Src` folder where each Subfolder represents a Library in a site. You can alos create a subfolder called `_catalogs` and 
add a subfolder to that one called `masterpage` if you want to sync files to the Master Page Gallery.
17. Add files as you want to these folders
18. Press Ctrl-Shift-B to Build and let Gulp and gulp-spsync upload the files to SharePoint

# Using Gulp watchers

You can use Gulp watchers (gulp-watch) to upload files as they are changed. 
The following `gulpfile.js` shows how to upload all files on build and then upload files incrementally when changed and saved.

You need to run `npm install gulp-watch` to install the Gulp watcher

```javascript
var gulp = require('gulp')
var sp = require('gulp-spsync')
var watch = require('gulp-watch')

var settings = {
			"client_id":"...",
			"client_secret":"...",
			"realm" : "",
			"site" : "https://contoso.sharepoint.com/sites/site",
			"verbose": "true"
		};
gulp.task('default', function() {
	return gulp.src('src/**/*.*')
		.pipe(watch('src/**/*.*'))
		.pipe(sp(settings))		
		.pipe(gulp.dest('build'))
})

```
# Setting metadata for files
If you're files require metadata to be set when they are uploaded, you can pass in a metadata options (update_metadata, files_metadata).

**Example:**
```javascript
var fileMetadata = [
    {
        name: 'Item_Minimal.js',
        metadata: {
            "__metadata": { type: "SP.Data.OData__x005f_catalogs_x002f_masterpageItem" },
            Title: 'Item Minimal Template (via GULP)',
            MasterPageDescription: 'This is a display template added via gulp.',
            ManagedPropertyMapping: "'Path','Title':'Title'",
            ContentTypeId: '0x0101002039C03B61C64EC4A04F5361F38510660500A0383064C59087438E649B7323C95AF6',
            DisplayTemplateLevel: 'Item',
            TargetControlType: {
                "__metadata": {
                    "type": "Collection(Edm.String)"
                },
                "results": [
                    "SearchResults",
                    "Content Web Parts"
                ]
            }
        }
    },
    {
        name: 'Control_Minimal.js',
        metadata: {
            "__metadata": { type: "SP.Data.OData__x005f_catalogs_x002f_masterpageItem" },
            Title: 'Control Minimal Template (via GULP)',
            MasterPageDescription: 'This is a display template added via gulp.',
            ContentTypeId: '0x0101002039C03B61C64EC4A04F5361F38510660500A0383064C59087438E649B7323C95AF6',
            DisplayTemplateLevel: 'Control',
            TargetControlType: {
                "__metadata": {
                    "type": "Collection(Edm.String)"
                },
                "results": [
                    "SearchResults",
                    "Content Web Parts"
                ]
            }
        }
    }
];

var settings = {
    "client_id":"...",
    "client_secret":"...",
    "realm" : "",
    "site" : "https://contoso.sharepoint.com/sites/site",
    "verbose": true,
    "update_metadata": true,
    "files_metadata": fileMetadata
};
```

# Publishing files
By setting the **publish** setting, you can specify to publish your files when they are uploaded to the site.

```json
var settings = {
    "client_id":"...",
    "client_secret":"...",
    "realm" : "",
    "site" : "https://contoso.sharepoint.com/sites/site",
    "verbose": true,
    "publish": true
};
```

# Using nested folders (new in 1.4.0)

If you're using nested folders or deep structures, you can choose the name of the "start folder", using the `startFolder` option. 
Assume you have your SharePoint files under `src/template1/_sp/_catalogs` and `src/template2/_sp/_catalogs/` then you can use `"startFolder"="_sp"` to make sure that the first folder names are stripped.

```javascript
var gulp = require('gulp')
var sp = require('gulp-spsync')

var settings = {
			"client_id":"...",
			"client_secret":"...",
			"realm" : "",
			"site" : "https://contoso.sharepoint.com/sites/site",
			"verbose": "true",
            "startFolder":"_sp"
		};
gulp.task('default', function() {
	return gulp.src('src/**/_sp/**/*.*')
		.pipe(sp(settings))		
})

```