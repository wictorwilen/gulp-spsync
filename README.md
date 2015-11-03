# gulp-spsync
Gulp plugin for synchronizing local files with a SharePoint library

# Features
 
* Gulp plugin
* Copies local files to a SharePoint Document library

# Planned features
* Support for sub folders
* Support for multiple libraries
 
# How to use

1. Prepare SharePoint by registering a SharePoint app using appregnew.aspx. Eg go to https://contoso.sharepoint.com/sites/site/_layouts/15/appregnew.aspx
2. Click on Generate for both Client Id and Client Secret. For Title, App Domain and Redirect URI, write something you don't care about. Then click on Create
3. Note down the Client Id and Client Secret, you will need it later
4. Navigate to appinv.aspx, https://contoso.sharepoint.com/sites/site/_layouts/15/appinv.aspx, enter the client ID in the App Id box and press Lookup
5. In the Permission Request XML text box enter the following XML and click Create  
```
<AppPermissionRequests AllowAppOnlyPolicy="true">
    <AppPermissionRequest
        Scope="http://sharepoint/content/sitecollection/web"
        Right="Manage"/>
</AppPermissionRequests>
```
6. In the following consent screen choose to trust the App by clicking on Trust It!
7. Open a folder using Visual studio code
8. Run `npm install gulp` to install the Gulp task runner
9. Run `npm install wictorwilen/gulp-spssync` to install to install the gulp-spsync 
10. Press Ctrl-Shift-P, type Task and choose to Configure Task Runner
11. In the tasks.json file that is being created replace the contents with the following:
```
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
12. Create a new file in the root of your folder called `gulpfile.js`, and modify it as follows
```
var gulp = require('gulp')
var sp = require('gulp-spsync')
gulp.task('default', function() {
return gulp.src('src/*.*').
    pipe(sp({
        "client_id":"3d271647-2e12-4ae5-9271-04b3aa67dcd3",
        "client_secret":"Zk9ORywN0gaGljrtlxfp+s5vh7ZyWV4dRpOXCLjtl8U=",
        "realm" : "",
        "site" : "https://contoso.sharepoint.com/sites/site",
        "library": "Site Assets",
        "verbose": "true"
    })).		
    pipe(gulp.dest('build'))
})
```
13. Replace the client_id and client_secret parameters with the value for the App you just created
14. Replace the site URL with your site URL and the library name with your library name
15. Create a folder called `Src`
16. Add any number of files to the Src folder
17. Press Ctrl-Shift-B to Build and let Gulp and gulp-spsync upload the files to SharePoint 

# Using Gulp watchers

You can use Gulp watchers (gulp-watch) to upload files as they are changed. 
The following `gulpfile.js` shows how to upload all files on build and then upload files incrementally when changed and saved.

You need to run `npm install gulp-watch` to install the Gulp watcher

```
var gulp = require('gulp')
var sp = require('gulp-spsync')
var watch = require('gulp-watch')

var settings = {
			"client_id":"...",
			"client_secret":"...",
			"realm" : "",
			"site" : "https://contoso.sharepoint.com/sites/site",
			"library": "Documents",
			"verbose": "true"
		};
gulp.task('default', function() {
	return gulp.src('src/*.*')
		.pipe(watch('src/*.*'))
		.pipe(sp(settings))		
		.pipe(gulp.dest('build'))
})

```


