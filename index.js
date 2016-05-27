'use strict'
var through = require('through2')
var rp = require('request-promise');
var u = require('url')
var gutil = require('gulp-util');
var path = require('path');
var util = require("util");

module.exports = function(args){
	var options = {
		client_id: "",
		client_secret: "",
		realm: "",
		site: "",
		verbose: false,
		watch: false,
		update_metadata: false,
		files_metadata: [],
		publish: false
	}
	
	if(!args){
		throw "options required"
	}
	if(!args.client_id){
		throw "The client_id options parameter is required"
	}
	if(!args.client_secret){
		throw "The client_secret options parameter is required"
	}
	if(!args.site){
		throw "The site options parameter is required"
	}
	
	if (args) {
		// Required properties
		options.client_id = args.client_id;
		options.client_secret = args.client_secret;
		options.site = args.site;
		// Default properties or configured via the gulp script
		options.realm = args.realm || options.realm;
		options.verbose = args.verbose || options.verbose;
		options.watch = args.watch || options.watch;
		options.update_metadata = args.update_metadata || options.update_metadata;
		options.files_metadata = args.files_metadata || options.files_metadata;
		options.publish = args.publish || options.publish;
	}
	
	var getFormattedPrincipal = function (principalName, hostName, realm){
		var resource = principalName
		if(hostName != null && hostName != "" ) {
			resource += "/" + hostName 	
		} 
		resource += "@" + realm
		return resource
	}
	
	var toDateFromEpoch = function(epoch){
  		var tmp = parseInt(epoch); 
		if(tmp<10000000000) tmp *= 1000;	
		var d = new Date()
		d.setTime(tmp)
		return d;
	}
	var now = function() {
		return new Date()
	}
	var globalEndPointPrefix = "accounts";
    var acsHostUrl = "accesscontrol.windows.net";
	var acsMetadataEndPointRelativeUrl = "/metadata/json/1";
	var S2SProtocol = "OAuth2"
	var sharePointPrincipal = "00000003-0000-0ff1-ce00-000000000000"
	var bearer = "Bearer realm=\""
	var https ="https://"
	var clientsvc = "/vti_bin/client.svc"
	var tokens = null
	
	var getStsUrl = function(realm){
		if(options.verbose){
			gutil.log('Locating STS Url for ' + realm)	
		}
		var url = https + globalEndPointPrefix + "." + acsHostUrl + acsMetadataEndPointRelativeUrl + "?realm=" + realm
		return rp
			.get(url)
			.then(function(data){
				var endpoints =JSON.parse(data).endpoints 
				for(var i in endpoints){
					if(endpoints[i].protocol == S2SProtocol  )
					{
						if(options.verbose){
							gutil.log('STS Endpoint found ' + endpoints[i].location)	
						}
						return endpoints[i].location
					}
				}	
				throw "ACS endpoint not found"
			});
	}
	var getRealmFromTargetUrl = function(targetUrl){
		if(options.verbose){
			gutil.log('Locating realm for ' + targetUrl)	
		}
		
		return rp.post( targetUrl + clientsvc,{
			headers: {
				"Authorization": "Bearer "
			},
			resolveWithFullResponse: true
		}).then(function(response){
			throw "Unexpected"
		}).catch(function(err){
			if(err.name== 'RequestError'){
				throw "Request error"
			}
			var headers = err.response.headers	
			var data = headers["www-authenticate"]
			var ix  = data.indexOf(bearer)	+ bearer.length
			data = data.substring(ix, ix+36)
			if(options.verbose){
				gutil.log('Realm is ' + data)	
			}
			return data; 
		});
	}
	var getAppOnlyAccessToken = function(
		targetPrincipalName,
		targetHost,
		targetRealm){
		
		var resource = getFormattedPrincipal(targetPrincipalName, targetHost, targetRealm)		
		var clientId = getFormattedPrincipal(options.client_id, "", targetRealm)
		
		var httpOptions = {
			headers: {
				"Content-Type": "application/x-www-form-urlencoded"
			},
			form: {
				"grant_type": "client_credentials",
				"client_id": clientId,
				"client_secret": options.client_secret,
				"resource": resource
			}
		};
		
		if(options.verbose){
			gutil.log('Retreiving access token for  ' + clientId)	
		}
		return getStsUrl(options.realm)
			.then(function(stsUrl){
				return rp.post(stsUrl, httpOptions)
					.then(function(data){
						return JSON.parse(data)		
					})
				});		
	}
	
	var updateFileMetadata = function(filename, library, tokens) {
		// Check if the filename is in the array with metadata
		var fileMetadata = options.files_metadata.filter(function (fm) {
			if (fm.name === filename) {
				return fm;
			}
		});
		// Check if metadata has been retrieved for the file
		if (fileMetadata.length > 0) {
			var metadata = fileMetadata[0].metadata;
			var metadataHeader = {
				"headers":{
					"Authorization": "Bearer " + tokens.access_token,
					"content-type":"application/json;odata=verbose",
					"X-HTTP-Method": "PATCH",
					"If-Match": "*"
				},
				"body": JSON.stringify(metadata) 
			};
			rp.post(
				options.site + "/_api/web/GetFolderByServerRelativeUrl('" + library +"')/Files('"+filename +"')/listitemallfields",
				metadataHeader
			).then(function(postData) {
				if (options.publish) {
					publishFile(filename, library, tokens);
				} else {
					gutil.log(gutil.colors.green('Uploaded and metadata added successful'));
					return postData;
				}
			}).catch(function(err){
				gutil.log(gutil.colors.red("Unable to update metadata of the file"));
			});
		} else {
			// Check if the file has to be published
			if (options.publish) {
				publishFile(filename, library, tokens);
			}
		}
	}
	
	var publishFile = function (filename, library, tokens) {
		var publishHeader = {
			"headers":{
				"Authorization": "Bearer " + tokens.access_token,
				"content-type":"application/json;odata=verbose"
			}
		};
		// First check out the file
		rp.post(
			options.site + "/_api/web/GetFolderByServerRelativeUrl('" + library +"')/Files('"+filename +"')/CheckOut()",
			publishHeader
		).then(function(result){
			// Check in major version
			rp.post(
				options.site + "/_api/web/GetFolderByServerRelativeUrl('" + library +"')/Files('"+filename +"')/CheckIn(comment='Checked in via GULP', checkintype=1)",
				publishHeader
			).then(function (result) {
				gutil.log(gutil.colors.green('Published file'));
				return result;
			})
		}).catch(function(err){
			gutil.log(gutil.colors.red("Unable to publish file"));
			gutil.log(gutil.colors.red(err));
		});
	}
	
	var uploadFile = function(file, content){
		var headers = {
			"headers":{
				"Authorization": "Bearer " + tokens.access_token,
				"content-type":"application/json;odata=verbose",
				"accept":"application/json;odata=verbose"
			},
			"body": content
		};
		
		var ix = file.relative.lastIndexOf(path.sep)
        var ix2 = 0;
        if(options.startFolder) {
            ix2 = file.relative.indexOf(options.startFolder) + options.startFolder.length + 1
            if(ix2 == -1) {
                ix2 = 0
            }
        }
		var library = file.relative.substring(ix2,ix)
        if(options.verbose){
            gutil.log('Using library: ' + library)	
        }
		var filename = file.relative.substring(ix+1)
		
		if(path.sep == "\\"){
			library = library.replace(/\\/g, "/")
		}
		
		return checkFoldersAndCreateIfNotExist(library, filename, options, tokens).then(function() {
			return rp.post(
				options.site + "/_api/web/GetFolderByServerRelativeUrl('" + 
				library +"')/Files/add(url='"+
				filename +"',overwrite=true)",
				headers
			)
			.then(function(success) {
				if (options.files_metadata.length > 0 && options.update_metadata) {
					updateFileMetadata(filename, library, tokens);
				} else {
					if (options.publish) {
						publishFile(filename, library, tokens);
					} else {
						gutil.log(gutil.colors.green('Upload successful'))	
						return success;
					}
				}
			})
			.catch(function(err){
				switch(err.statusCode){
					case 423:
						gutil.log(gutil.colors.red("Unable to upload file, it might be checked out to someone"))
						break;
					default:
						gutil.log(gutil.colors.red("Unable to upload file, it might be checked out to someone"))
						break;
				}
			});
		});
	}
	
	var checkFoldersAndCreateIfNotExist = function(library, filename, options, tokens) {
		var foldersArray = getFolderPathsArray(library);
		var proms = [];
		foldersArray.forEach(function (val, index) {
			proms.push(checkFolderExists(val));
		});
		
		return Promise.all(proms)
			.then(function(data) {
				var erroredIndexes = data.map(function (val, index) {
				if (val.error) {
					return index;
				}
			}).filter(function (x) { return x != undefined });
			var pathArray = [];
			erroredIndexes.forEach(function (val, index) {
				var path = foldersArray[val];
				pathArray.push(path);
			})
			if (pathArray.length > 0) {
				return createPathRecursive(pathArray, library, filename, options, tokens);
			}
		});
		
	}
	
	var checkFolderExists = function(folderName) {
		var getFolderUrl = util.format("/_api/web/GetFolderByServerRelativeUrl(@FolderName)" +
			"?@FolderName='%s'", encodeURIComponent(folderName));
        var opts = {
			headers: {
				"Accept": "application/json;odata=verbose",
                "Authorization": "Bearer " + tokens.access_token,
                "content-type":"application/json;odata=verbose"
			},
			json: true
		};
		var endPoint = options.site + getFolderUrl;
		if(options.verbose){
			gutil.log("Checking folder exists " + endPoint);
		}
		return rp.get(endPoint, opts)
			.then(function (success) {
				if(options.verbose){
					gutil.log('Folder ' + folderName + ' exists');
				}
				return success;
			})
			.catch(function(err) {
				gutil.log("INFO: Folder '" + folderName + "' doesn't exist and will be created");
				return err;
			});
	}
	
	var createPathRecursive = function(path, library, filename, options, tokens) {
		if(options.verbose){
			gutil.log("Creating path " + path[0]);
		}
		var setFolder = util.format("/_api/web/folders");
		var body = "{'__metadata': {'type': 'SP.Folder'}, 'ServerRelativeUrl': '" + path[0] + "'}";
		var opts = {
			headers: {
				"Accept": "application/json;odata=verbose",
				"Authorization": "Bearer " + tokens.access_token,
				"content-type": "application/json;odata=verbose",
				"content-length": Buffer.byteLength(body)
			},
			body: body
		};
				  
		return new Promise(function (resolve) {
			rp.post(options.site + setFolder, opts)
			.then(function (res) {
				resolve(path.slice(1, path.length));
			})
			.catch(function(err) {
				gutil.log("ERR: " + err);
				return err;
			});
		})
		.then(function (path) {
			if (path.length > 0) {
				return createPathRecursive(path, library, filename, options, tokens);
			}
			return true;
		});		
	}
	
	var getFolderPathsArray = function (folder) {
		if (endsWith(folder, "/") && folder !== "/") {
			folder = folder.slice(0, -1);
		}

		var folderNamesArray = folder.split('/');
		var foldersArray = [];
		for (var i = 0; i < folderNamesArray.length; i++) {
			var pathArray = [];
			for (var r = 0; r <= i; r++) {
				pathArray.push(folderNamesArray[r]);
			}
			foldersArray.push(pathArray.join('/'));
		}
		return foldersArray;
	}

	var endsWith = function(str, suffix) {
		return str.indexOf(suffix, str.length - suffix.length) !== -1;
	}

	return through.obj(function(file, enc, cb){
		// If watch is set to true, only upload the changed files
		if (options.watch && file.event !== "change") {
			if(options.verbose){
				gutil.log("Skipping:", gutil.colors.yellow(file.relative))
			}
			cb(null,file)
			return;
		}
		
		var fileDone = function(parameter) {
			cb(null,file)
		}
		
		if(file.isNull()){
			cb(null, file)
			return;
		}
		if (file.isStream()) { 
 			cb(new gutil.PluginError("gulp-spsync", 'Streaming not supported')); 
			return; 
		} 

		var content = file.contents; 
        if (file.contents == null || file.contents.length === 0) { 
             content = ''; 
        } 

		gutil.log('Uploading ' + file.relative)
				
		if(tokens == null || now() > toDateFromEpoch(tokens.expires_on) ){
			getRealmFromTargetUrl(options.site).then(function(realm){
				return getAppOnlyAccessToken(
					sharePointPrincipal,
					u.parse(options.site).hostname,
					realm)
					.then(function(token){
						tokens = token
						return uploadFile(file, content).then(fileDone)
					})
			}).catch(function(err){
				cb(new gutil.PluginError("gulp-spsync", err)); 
			});	
		} else {
			return uploadFile(file, content).then(fileDone)
		}
	},function(cb){
		if(options.verbose){
			gutil.log("And we're done...")	
		}		
		cb();
	})
}