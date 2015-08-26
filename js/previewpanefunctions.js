/* This file is used by previewpaane Menu */
//Register any required scripts
RegisterSod('sp.requestexecutor.js', '/_layouts/15/sp.requestexecutor.js');

var LoadAndExecuteSodFunction = window.LoadAndExecuteSodFunction || function (scriptKey, fn) {
    if (!ExecuteOrDelayUntilScriptLoaded(fn, scriptKey)) {
        LoadSodByKey(NormalizeSodKey(scriptKey));
    }
}



window.SPOSearchCustomizations = window.SPOSearchCustomizations || {};
window.SPOSearchCustomizations.Automation = window.SPOSearchCustomizations.Automation || {};
window.SPOSearchCustomizations.Automation.PreviewPaneConstants = window.SPOSearchCustomizations.Automation.PreviewPaneConstants || {
    viewHistoryLink: 'viewHistory',
    noVersioningMsg: 'noVersioningMsg',
    addToOneDriveLink: 'addToOneDriveLink',
    docViewEdit: 'DocViewEdit',
    docFollowStatusLink: 'DocFollowStatusLink',
    docShareLink: 'DocShareLink',
    ctxMenuLink: 'ctxMenuLink',
    ctxMenuContainer: 'ctxMenuContainer',
    addToOneDriveLink: 'AddToOneDriveLink',
    versionHistoryLink: 'VersionHistoryLink',
    contactNameLink: 'ContactNameLink',
    viewPropsLink: 'ViewProps',
    editPropsLink: 'EditProps',
    setAlertsLink: 'SetAlerts',
    checkInLink: 'CheckIn',
    checkOutLink: 'CheckOut',
    discardCheckOutLink: 'DisscardCheckout',
    publishLink: 'Publish',
    downloadCopyLink: 'DownloadCopy',
    deleteFileLink: 'deleteFile',
    uploadNewVersionLink: 'uploadNewVersion'
};


var ppConstants = window.SPOSearchCustomizations.Automation.PreviewPaneConstants;

window.SPOSearchCustomizations.Automation.PreviewPaneDataOperations = function () {
    var ConcatenateMetaDataValues = function (brandvalue) {
        var processedEntities = [];
        if (brandvalue != null) {
            var tokens = brandvalue.split(';');

            for (var i = 0; i < tokens.length; i++) {

                var subTokens = tokens[i].split('|');
                if (subTokens.length == 2 | subTokens.length == 1)
                    processedEntities.push(subTokens[0]);
                else if (subTokens.length == 3)
                    processedEntities.push(subTokens[2]);
            }
        }
        return processedEntities.join()
            .replace(',', ', ');
    }
    return {
        ConcatenateMetaDataValues: ConcatenateMetaDataValues
    }
}();

window.SPOSearchCustomizations.Automation.PreviewPaneOperations = window.SPOSearchCustomizations.Automation.PreviewPaneOperations || function () {

    var LoadSiteUrlFromFullUrl = function (docurl) {
        return $.Deferred(function (def) {
            LoadAndExecuteSodFunction('sp.requestexecutor.js', function () {
                var executor = new SP.RequestExecutor(_spPageContextInfo.webAbsoluteUrl);
                var info = {
                    url: docurl.substring(0, docurl.lastIndexOf('/')) + '/_api/contextinfo',
                    method: "POST",
                    headers: {
                        'Accept': 'application/json; odata=verbose',
                        'content-type': 'application/json;odata=verbose'
                    },
                    success: function (data) {
                        var body = JSON.parse(data.body);
                        if (body.d) {
                            var webUrl = body.d.GetContextWebInformation.WebFullUrl;
                            def.resolve(webUrl);
                        } else {
                            def.reject('no weburl found')
                        }
                    },
                    error: function (s, a, errMsg) {
                        def.reject(errMsg)
                    }
                };
                executor.executeAsync(info);
            })
        });
    };

    var LoadListAndItemData = function (weburl, listId, docPath, itemId) {
        return $.Deferred(function (def) {
            LoadAndExecuteSodFunction('sp.js', function () {
                var clientContext = new SP.ClientContext(weburl);
                var web = clientContext.get_web();
                var list = clientContext.get_web()
                    .get_lists()
                    .getById(listId);
                var listItem = list.getItemById(itemId);
                clientContext.load(web);
                clientContext.load(list);

                clientContext.load(listItem, 'EffectiveBasePermissions');
                clientContext.executeQueryAsync(function () {
                    //window.list = list.get_defaultDisplayFormUrl();				
                    def.resolve({
                        web: web,
                        list: list,
                        listItem: listItem
                    })
                }, function (s, a) {
                    def.reject(a.get_message());
                });
            });
        });
    };
    var LoadUser = function (username, weburl, csrid) {
        if (username) {
            var userid = username.split(';')[1];
            if (userid) {
                $('#' + csrid + ppConstants.contactNameLink)
                    .attr('href', weburl + '/_layouts/15/userdisp.aspx?id=' + userid);
            } else
                return $.Deferred(function (def) {
                    LoadAndExecuteSodFunction('sp.js', function () {
                        var clientContext = new SP.ClientContext(weburl);
                        var user = clientContext.get_web()
                            .ensureUser(username);
                        clientContext.load(user);
                        clientContext.executeQueryAsync(function () {
                            //window.list = list.get_defaultDisplayFormUrl();				
                            def.resolve(user)
                        }, function (s, a) {
                            def.reject(a.get_message());
                        });
                    });
                });
        }
        //end if - do nothing if username is null
    };



    var LoadIsFollowing = function (docPath, weburl) {
        return $.Deferred(function (def) {
            LoadAndExecuteSodFunction("sp.requestexecutor.js", function () {
                var requestinfo = {
                    url: weburl + '/_api/social.following/isfollowed',
                    method: "POST",
                    body: JSON.stringify({
                        "actor": {
                            "__metadata": {
                                "type": "SP.Social.SocialActorInfo"
                            },
                            "ActorType": SP.Social.SocialActorType.document,
                            "ContentUri": docPath,
                            "Id": null
                        }
                    }),
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-type": "application/json;odata=verbose"
                    },
                    success: function (responseData) {
                        var jsonObject = JSON.parse(responseData.body);
                        def.resolve(jsonObject.d.IsFollowed);
                    },
                    error: function (s, a, errMsg) {
                        def.reject(errMsg);
                    }
                };
                var executor = new SP.RequestExecutor(weburl);
                executor.executeAsync(requestinfo);
            });
        });
    };

    var ToggleFollowUnfollow = function (weburl, docPath, isAlreadyFollowing) {
        var toggleurl;
        return $.Deferred(function (def) {
            toggleurl = isAlreadyFollowing ? _spPageContextInfo.webAbsoluteUrl + '/_api/social.following/stopfollowing' : _spPageContextInfo.webAbsoluteUrl + '/_api/social.following/follow';

            var requestinfo = {
                url: toggleurl,
                method: "POST",
                body: JSON.stringify({
                    "actor": {
                        "__metadata": {
                            "type": "SP.Social.SocialActorInfo"
                        },
                        "ActorType": SP.Social.SocialActorType.document,
                        "ContentUri": docPath,
                        "Id": null
                    }
                }),
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-type": "application/json;odata=verbose"
                },
                success: function (responseData) {

                    var jsonObject = JSON.parse(responseData.body);
                    var isfollowed = jsonObject.d.Follow;
                    var stoppedFollowing = jsonObject.d.StopFollowing;

                    def.resolve(!isAlreadyFollowing)
                },
                error: function (s, a, errMsg) {
                    def.reject(errMsg);
                }
            };
            var executor = new SP.RequestExecutor(weburl);
            executor.executeAsync(requestinfo);

        });
    };

    var initPreviewPaneData = function (props) {

        // bind non async data first
        //Share ,versionhistory and AddToOneDrive Links.
        var shareurl = props.webUrl + '/' + _spPageContextInfo.layoutsUrl + '/aclinv.aspx?forSharing=1&List=' + encodeURIComponent(props.listId) + '&obj=' + encodeURIComponent(props.listId) + ',' + encodeURIComponent(props.itemId) + ',DOCUMENT';
        var versionhistoryurl = props.webUrl + '/' + _spPageContextInfo.layoutsUrl + '/Versions.aspx?list=' + encodeURIComponent(props.listId) + '&Id=' + encodeURIComponent(props.itemId);
        var loginName = _spPageContextInfo.userLoginName;
        var mysiteurl = 'https://domain-my.sharepoint.com'.toLowerCase().split(":443")[0] + '/personal/' + loginName.replace(/\./g, '_').replace(/\@/, '_') + '/Documents/';
        var addToOneDriveUrl = props.webUrl + '/' + _spPageContextInfo.layoutsUrl + '/copy.aspx?SourceUrl=' + props.docPath + '&FldUrl=' + mysiteurl;
        //bind share event
        $('#' + props.csrid + ppConstants.docShareLink)
            .on('click', function (e) {
                e.preventDefault ? e.preventDefault() : e.returnValue = false;
                OpenPopUpPage(shareurl);
                return false;
            });
        //bind version history event
        $('#' + props.csrid + ppConstants.versionHistoryLink)
            .on('click', function (e) {
                e.preventDefault ? e.preventDefault() : e.returnValue = false;
                OpenPopUpPage(versionhistoryurl);
                return false;
            });
        //Bind add To OneDrive Url
        $('#' + props.csrid + ppConstants.addToOneDriveLink)
            .on('click', function (e) {
                e.preventDefault ? e.preventDefault() : e.returnValue = false;
                OpenPopUpPage(addToOneDriveUrl);
                return false;
            });

        //Load List item data and other async calls to bind data

        $.when(LoadListAndItemData(props.webUrl, props.listId, props.docPath, props.itemId))
            .done(function (data) {
                var viewEditlink = $('#' + props.csrid + ppConstants.docViewEdit);
                var downloadLink = $('#' + props.csrid + ppConstants.downloadCopyLink + ' a');
                if (!(viewEditlink.hasClass('pdfhref'))) {
                    if (data.listItem.get_effectiveBasePermissions()
                        .has(SP.PermissionKind.editListItems)) {
                        viewEditlink.text('Edit').attr('href', props.serverRedirectUrl ? props.serverRedirectUrl.replace('action=default', 'action=edit') : '');
                    } else if (data.listItem.get_effectiveBasePermissions()
                        .has(SP.PermissionKind.viewListItems)) {
                        viewEditlink.text('View')
                            .attr('href', props.serverRedirectUrl);
                    }
                }

                // remove the version event if no versioning is enabled on the library
                if (!(data.list.get_enableVersioning())) {
                    $('#' + props.csrid + ppConstants.versionHistoryLink)
                        .off('click');
                    $('#' + props.csrid + ppConstants.versionHistoryLink)
                        .on('click', function (e) {
                            e.preventDefault ? e.preventDefault() : e.returnValue = false;
                            $('#' + props.csrid + ppConstants.noVersioningMsg)
                                .fadeIn('slow', function () {
                                    $(this)
                                        .fadeOut(2000);
                                })
                            return false;
                        });
                }
            })
            .fail(function (errMsg) {
                //alert(errMsg);
            });

        //Load ContactName
        var _defContactLink = LoadUser(props.contactName, props.webUrl, props.csrid);

        if (_defContactLink) {
            $.when(_defContactLink)
                .done(function (user) {
                    $('#' + props.csrid + ppConstants.contactNameLink)
                        .attr('href', props.webUrl + '/_layouts/15/userdisp.aspx?id=' + user.get_id());
                })
                .fail(function (errmsg) {
                    // alert(errmsg) // show err msg
                });
        }

        //Load Follow Status
        $.when(LoadIsFollowing(props.docPath, props.webUrl))
            .done(function (isFollowed) {
                var linktext = isFollowed ? "Unfollow" : "Follow";
                $('#' + props.csrid + ppConstants.docFollowStatusLink)
                    .attr('data-isfollowed', isFollowed)
                    .text(linktext)
                    .click(function (e) {
                        var _telement = $(this);
                        e.preventDefault ? e.preventDefault() : e.returnValue = false;
                        $.when(LoadIsFollowing(props.docPath, props.webUrl))
                            .done(function (isFollowed) {
                                $.when(ToggleFollowUnfollow(props.webUrl, props.docPath, isFollowed))
                                    .done(function (isFollowed) {
                                        var linktext = isFollowed ? "Unfollow" : "Follow"
                                        _telement.attr('data-isfollowed', isFollowed)
                                            .text(linktext);
                                    })
                                    .fail(function (errMsg) {
                                        //If Follwing/Unfollowing throws error, then call OOTB method to Show Error.
                                        if (HP && HP.Follow) {
                                            HP.Follow(props.docPath, !isFollowed, true)
                                        }
                                    });
                            })
                            .fail(function (errMsg) {
                                //alert(errMsg)
                            })
                    });
            })
            .fail(function (errMsg) {
                //if status loading is failed then attach OOTB follow functionality to Show Error on UI
                if (HP && HP.Follow) {
                    $('#' + props.csrid + ppConstants.docFollowStatusLink)
                        .attr('data-isfollowed', false)
                        .text('fail')
                        .click(function (e) {
                            var _telement = $(this);
                            e.preventDefault ? e.preventDefault() : e.returnValue = false;
                            HP.Follow(props.docPath, true, true);
                            return false;
                        });
                }
            })
    };
    return {
        initPreviewPaneData: initPreviewPaneData
    }
}();


// Code for "..." menu

window.SPOSearchCustomizations.Automation.PreviewPaneOperations.ContextMenu = window.SPOSearchCustomizations.Automation.PreviewPaneOperations.ContextMenu || function () {

    var checkout = function (props, callback) {
        props.file.checkOut();
        props.context.executeQueryAsync(function (data) {
            if (callback) callback();
        }, function (s, a) {
            //alert(a.get_message())
        });
    };
    var checkIn = function (props, callback) {
        OpenPopUpPage(props.webUrl + '/' + _spPageContextInfo.layoutsUrl + '/checkin.aspx?List=' + props.listId + '&FileName=' + encodeURIComponent(props.file.get_serverRelativeUrl()),
            function (status, data) {
                if (status === SP.UI.DialogResult.OK) {
                    if (callback) {
                        callback()
                    }
                }
            }, null, null);
    };

    var DiscardCheckout = function (props, callback) {
        window.fff = props.file;
        props.file.undoCheckOut();
        props.context.executeQueryAsync(function () {
            if (callback) {
                callback();
            }
        },
            function (s, a) {
                //alert(a.get_message())
            });
    }

    var registerViewPropsClick = function (props) {
        $('#' + props.csrid + ppConstants.viewPropsLink)
            .on('click', function (e) {
                OpenPopUpPage(props.webUrl + '/' + _spPageContextInfo.layoutsUrl + '/listform.aspx?PageType=' + SP.PageType.displayForm + '&ListId=' + props.listId + '&ID=' + props.itemId);
                return false;
            })
    };

    var registerCheckoutLink = function (props) {
        $('#' + props.csrid + ppConstants.checkOutLink)
            .on('click', function (e) {
                checkout(props, function () { });
            });
    };

    var registercheckInLink = function (props) {
        $('#' + props.csrid + ppConstants.checkInLink)
            .on('click', function (e) {
                checkIn(props, function () { })
            });
    };

    var registerDiscardCheckoutLink = function (props) {
        $('#' + props.csrid + ppConstants.discardCheckOutLink)
            .on('click', function (e) {
                DiscardCheckout(props, function () { });
            });
    };

    // upload file methods
    var createDialogElement = function (elementid, extention) {
        var dialogElement = $($('#UploadNewVersionDialogMarkup')
            .html());
        var span = dialogElement.find('.file-ext-note')
            .html('upload a file with ' + extention + ' extension');
        return {
            statusspan: span,
            element: dialogElement.get(0),
            uploadBtn: dialogElement.find('.uploadNewVersionBtn'),
            newFilePicker: dialogElement.find('.newfilepicker')
        };
    };

    var editTitle = function (props, callback) {
        props.waitscreen = SP.UI.ModalDialog.showWaitScreenWithNoClose('Working on it', 'Please Wait');
        var title = props.listItem.get_fieldValues().Title;
        var description = props.listItem.get_fieldValues().AppendingDescription;
        var contenttypeid = props.listItem.get_fieldValues().ContentTypeId;
        props.newItem = props.list.getItemById(props.itemId);
        props.context.load(props.newItem);
        props.context.executeQueryAsync(function () {
            props.newItem.set_item('Title', title);
            props.newItem.set_item('AppendingDescription', description);
            props.newItem.set_item('ContentTypeId', contenttypeid);
            props.newItem.update();
            props.context.executeQueryAsync(function () {
                callback();
                props.waitscreen.close();
            }, function (s, a) {
                props.waitscreen.close();
                SP.UI.ModalDialog.showErrorDialog('Error', a.get_message());
            });
        }, function (s, a) {
            props.waitscreen.close();
            SP.UI.ModalDialog.showErrorDialog('Error', a.get_message());
        });
    };
    var doUpload = function (props) {
        //checkout file and upload

        var doUploadAction = function () {
            var executor = new SP.RequestExecutor(props.webUrl);
            var info = {
                headers: {
                    "Accept": "application/json; odata=verbose",
                    "content-type": "application/json;odata=verbose",
                },
                url: props.webUrl + "/_api/web/lists(guid'" + props.listId.slice(1, -1) + "')/RootFolder/Files/Add(url=@filename,overwrite=true)?@filename='" + encodeURIComponent(props.file.get_name()) + "'",
                method: "POST",
                binaryStringRequestBody: true,
                body: props.body,
                success: function (data) {
                    SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.OK, {
                        listid: props.listId,
                        itemid: props.itemId,
                        filename: props.file.get_name()
                    });
                    //delete props extra values
                },
                error: function (sender, args, errMsg) {
                    props.popupitem.statusspan.text(errMsg)
                        .addClass('ms-error')
                        .fadeOut()
                        .fadeIn();
                    $('#ms-aclinv-prgBar')
                        .hide();
                }
            };
            executor.executeAsync(info);
        };

        if (props.checkedOutBy.get_serverObjectIsNull() || props.file.get_checkOutType() === SP.CheckOutType.none) {
            checkout(props, doUploadAction);
        } else {
            doUploadAction();
        }
    };
    var _arrayBufferToBase64 = function (buffer) {
        var binary = '';
        var bytes = new window.Uint8Array(buffer);
        var len = bytes.byteLength;
        for (var i = 0; i < len; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return binary;
    };

    var showUploadDialog = function (props) {
        var oldname = props.file.get_name();
        var oldext = oldname.substring(oldname.lastIndexOf('.'));
        var popupitem = createDialogElement('UploadNewVersionDialogMarkup', oldext);
        var options = SP.UI.$create_DialogOptions();
        options.html = popupitem.element;
        options.title = "Upload New Version";
        options.allowMaximize = false;
        options.showClose = true;
        options.autoSize = true;
        options.dialogReturnValueCallback = function (sts, data) {
            if (sts === SP.UI.DialogResult.OK) {
                editTitle(props, function () {
                    OpenPopUpPage(props.webUrl + '/_layouts/15/listform.aspx?PageType=6&ListId=' + props.listId + '&ID=' + props.itemId + '&Mode=Upload'); //&Mode=Upload
                });
            } else {

            }
        };
        SP.UI.ModalDialog.showModalDialog(options);

        //register change event for file picker
        popupitem.newFilePicker.bind('change', function (evt) {
            var files = evt.target.files;
            window.filesss = files;
            if (files.length === 1) {
                //popupitem.statusspan.text('please upload a file').addClass('ms-error');;
                var file = files[0];
                var newname = file.name;
                var oldname = props.file.get_name();
                var oldext = oldname.substring(oldname.lastIndexOf('.'));
                var newext = newname.substring(newname.lastIndexOf('.'));
                if (oldext.toLowerCase() !== newext.toLowerCase()) {
                    popupitem.statusspan.removeClass('file-success').addClass('ms-error').html('upload a file with ' + oldext.toLowerCase() + ' extension')
                        .fadeOut()
                        .fadeIn();
                } else {

                    popupitem.statusspan.removeClass('ms-error').addClass('file-success').html('&#10004;')
                        .fadeOut()
                        .fadeIn();
                }
            }
            else {
                var oldname = props.file.get_name();
                var oldext = oldname.substring(oldname.lastIndexOf('.'));
                popupitem.statusspan.removeClass('file-success').addClass('ms-error').html('upload a file with ' + oldext.toLowerCase() + ' extension')
                        .fadeOut()
                        .fadeIn();
            }
        });
        //register click event
        popupitem.uploadBtn.click(function () {
            var files = popupitem.newFilePicker.get(0)
                .files;
            if (files.length === 1) {
                //upload doUpload()
                var file = files[0];

                var newname = file.name;
                var oldname = props.file.get_name();
                var oldext = oldname.substring(oldname.lastIndexOf('.'));
                var newext = newname.substring(newname.lastIndexOf('.'));
                if (oldext.toLowerCase() === newext.toLowerCase()) {
                    var reader = new FileReader();
                    reader.onload = (function (theFile) {
                        return function (e) {
                            var body = null;
                            if (FileReader.prototype.readAsBinaryString)
                                body = e.target.result;
                            else
                                body = _arrayBufferToBase64(e.target.result);
                            //uploadFile();
                            props.newFile = theFile;
                            props.body = body;
                            props.popupitem = popupitem;
                            doUpload(props);
                        };
                    })(file);
                    if (reader.readAsBinaryString)
                        reader.readAsBinaryString(file);
                    else
                        reader.readAsArrayBuffer(file);

                    $('#ms-aclinv-prgBar')
                        .show();
                } else {
                    popupitem.statusspan.addClass('ms-error')
                        .fadeOut()
                        .fadeIn();
                }
            } else {
                popupitem.statusspan.addClass('ms-error')
                    .fadeOut()
                    .fadeIn();
            }
            //upload document to library
        });
    };


    var registerEditPropsClick = function (props) {
        $('#' + props.csrid + ppConstants.editPropsLink)
            .on('click', function (e) {
                var hasApprove = props.hasApprovePermission;
                var isCheckoutRequired = props.isCheckoutRequired;
                var isCurrentusercheckedout = props.CurrentUserCheckedout;

                if (isCheckoutRequired) {

                    if (props.checkedOutBy.get_serverObjectIsNull() || props.CheckOutType === SP.CheckOutType.none) {
                        var result = confirm("This file is not checked out to you, You want to check out this file ?");
                        if (result == true) {
                            checkout(props, function () {
                                OpenPopUpPage(props.webUrl + '/' + _spPageContextInfo.layoutsUrl + '/listform.aspx?PageType=' + SP.PageType.editForm + '&ListId=' + props.listId + '&ID=' + props.itemId);
                            });
                        }
                    } else if (isCurrentusercheckedout) {
                        OpenPopUpPage(props.webUrl + '/' + _spPageContextInfo.layoutsUrl + '/listform.aspx?PageType=' + SP.PageType.editForm + '&ListId=' + props.listId + '&ID=' + props.itemId);
                    } else {
                        // alert('This file cannot be edited as it is currently checked out to ' + props.checkedOutBy.get_title());
                    }
                } else {
                    OpenPopUpPage(props.webUrl + '/' + _spPageContextInfo.layoutsUrl + '/listform.aspx?PageType=' + SP.PageType.editForm + '&ListId=' + props.listId + '&ID=' + props.itemId);
                }
                return false;
            });
    };


    var registersetAlertsLink = function (props) {
        $('#' + props.csrid + ppConstants.setAlertsLink)
            .on('click', function (e) {
                OpenPopUpPage(props.webUrl + '/' + _spPageContextInfo.layoutsUrl + '/SubNew.aspx?List=' + props.listId + '&ID=' + props.itemId);
                return false;
            })
    };
    var registerDeleteLink = function (props) {
        $('#' + props.csrid + ppConstants.deleteFileLink)
            .on('click', function (e) {
                var filename = props.file.get_name();
                if (window.confirm('Delete "' + filename + '" ??')) {
                    props.listItem.deleteObject();
                    props.context.executeQueryAsync(function (data) { }, function (s, a) { });
                }
                return false;
            });
    };
    var showElement = function (id) {
        $get(id)
            .style.display = "inherit";
    };

    var ShowUploadLink = function (id) {
        if (window.File && window.FileReader && window.FileList && window.Blob)
            $get(id)
            .style.display = "inherit";
    };

    var registerUploadLink = function (props) {
        $('#' + props.csrid + ppConstants.uploadNewVersionLink)
            .on('click', function (e) {
                var hasApprove = props.hasApprovePermission;
                var isCheckoutRequired = props.isCheckoutRequired;
                var isCurrentusercheckedout = props.CurrentUserCheckedout;

                if (isCheckoutRequired) {

                    if (props.checkedOutBy.get_serverObjectIsNull() || props.CheckOutType === SP.CheckOutType.none) {
                        var result = confirm("This file is not checked out to you, You want to check out this file ?");
                        if (result == true) {
                            checkout(props, function () {
                                showUploadDialog(props);
                            });
                        }
                    } else if (isCurrentusercheckedout) {
                        showUploadDialog(props);
                    } else {
                        //alert('This file cannot be edited as it is currently checked out to ' + props.checkedOutBy.get_title());
                    }
                } else {
                    showUploadDialog(props);
                }
                return false;
            })
    };
    var renderLinks = function (props) {
        if (props.hasViewPermission) {
            showElement(props.csrid + ppConstants.viewPropsLink);
            registerViewPropsClick(props);
        }
        if (props.hasEditPermission) {
            showElement(props.csrid + ppConstants.editPropsLink);
            registerEditPropsClick(props);

            ShowUploadLink(props.csrid + ppConstants.uploadNewVersionLink);
            registerUploadLink(props);
        }
        if (props.hasCreateAlertPermission) {
            showElement(props.csrid + ppConstants.setAlertsLink);
            registersetAlertsLink(props);
        }
        //no condition for download a copy
        showElement(props.csrid + ppConstants.downloadCopyLink);
        var downloadLink = props.webUrl + '/_layouts/15/download.aspx?SourceUrl=' + encodeURIComponent(props.docPath);

        $('#' + props.csrid + ppConstants.downloadCopyLink + ' a').attr('href', downloadLink); //update docPath if recently updated

        if (props.hasdeletePermission) {
            showElement(props.csrid + ppConstants.deleteFileLink);
            registerDeleteLink(props);
        }

        //logic to show checkin,checkout,publish
        if (props.hasEditPermission) {
            if (props.hasApprovePermission) {
                if (props.checkedOutBy.get_serverObjectIsNull() || props.CheckOutType === SP.CheckOutType.none) {
                    //show checkout
                    showElement(props.csrid + ppConstants.checkOutLink);
                    registerCheckoutLink(props);
                } else if (props.CheckOutType === SP.CheckOutType.online) //checkout online
                {
                    //show checkin and discardcheckout
                    showElement(props.csrid + ppConstants.checkInLink);
                    showElement(props.csrid + ppConstants.discardCheckOutLink);
                    registercheckInLink(props);
                    registerDiscardCheckoutLink(props);
                } else {
                    //do not do anthing if editing via client app
                    //showElement(props.csrid+ppConstants.deleteFileLink);
                    //registerDeleteLink(props);
                } //checkout using office client (ms word, powerpoint , excel)
            } else // no approve permissions
            {
                if (props.checkedOutBy.get_serverObjectIsNull() || props.CheckOutType === SP.CheckOutType.none) {
                    showElement(props.csrid + ppConstants.checkOutLink);
                    registerCheckoutLink(props);
                }
                    // checkout to current user
                else if (props.CurrentUserCheckedout) {
                    showElement(props.csrid + ppConstants.checkInLink);
                    showElement(props.csrid + ppConstants.discardCheckOutLink);
                    registercheckInLink(props);
                    registerDiscardCheckoutLink(props);
                } else // checked out another user
                {
                    showElement(props.csrid + ppConstants.checkOutLink);
                    registerCheckoutLink(props);

                }
            }
        }
        $('#' + props.csrid + ppConstants.ctxMenuContainer)
            .show();
    };

    var checkoutUserLoaded = function (props) {
        props.isCheckoutRequired = props.list.get_forceCheckout();
        props.moderationStatus = props.listItem.get_fieldValues()
            ._ModerationStatus;
        var permissions = props.listItem.get_effectiveBasePermissions();
        if (permissions.has(SP.PermissionKind.createAlerts)) {
            props.hasCreateAlertPermission = true;
        }
        if (permissions.has(SP.PermissionKind.deleteListItems)) {
            props.hasdeletePermission = true;
        }
        if (permissions.has(SP.PermissionKind.editListItems)) {
            props.hasEditPermission = true;
        }
        if (permissions.has(SP.PermissionKind.viewListItems)) {
            props.hasViewPermission = true;
        }
        if (permissions.has(SP.PermissionKind.approveItems)) {
            props.hasApprovePermission = true;
        }
        if (permissions.has(SP.PermissionKind.cancelCheckout)) {
            props.hasCancelCheckoutPermission = true;
        }

        props.CheckOutType = props.file.get_checkOutType()
        //}
        //checked out to current user
        if (props.checkedOutBy.get_serverObjectIsNull() || props.CheckOutType === SP.CheckOutType.none) {
            //do something
        } else if (props.checkedOutBy.get_loginName() === props.currentUser.get_loginName()) {
            props.CurrentUserCheckedout = true;
        }
            //checked to other user
        else {
            props.CurrentUserCheckedout = false;
        }
        renderLinks(props);


    };
    var checkForCheckout = function (props) {
        props.context.load(props.checkedOutBy);
        props.context.load(props.listItem, '_ModerationStatus', 'EffectiveBasePermissions', 'Title', 'AppendingDescription', 'ContentTypeId');
        props.context.load(props.file);
        props.context.load(props.list);
        props.context.load(props.currentUser);

        props.context.executeQueryAsync(
            function (data) {
                props.callBack(props);
            },
            function (s, a) {
                // alert(a.get_message());
            });
    };

    var init = function (props) {
        LoadAndExecuteSodFunction('sp.js', function () {
            props.ModerationStatusType = {
                Approved: 0,
                Denied: 1,
                Pending: 2,
                Draft: 3,
                Scheduled: 4
            };
            props.context = new SP.ClientContext(props.webUrl);
            props.web = props.context.get_web();
            props.currentUser = props.web.get_currentUser();
            props.list = props.web.get_lists()
                .getById(props.listId);
            props.listItem = props.list.getItemById(props.itemId);
            props.file = props.listItem.get_file();
            props.checkedOutBy = props.file.get_checkedOutByUser();
            props.callBack = checkoutUserLoaded;
            $('#' + props.csrid + ppConstants.ctxMenuLink)
                .off('click contextmenu');
            $('#' + props.csrid + ppConstants.ctxMenuLink)
                .on('click contextmenu', function (e) {
                    $('#' + props.csrid + ppConstants.ctxMenuContainer + ' li')
                        .hide()
                        .off('click');

                    $('#' + props.csrid + ppConstants.ctxMenuContainer + ' li')
                        .on('click', function (e) {
                            $('#' + props.csrid + ppConstants.ctxMenuContainer)
                                .hide();
                            //return false;
                        });
                    if ($('#' + props.csrid + ppConstants.ctxMenuContainer)
                        .is(":visible")) {
                        $('#' + props.csrid + ppConstants.ctxMenuContainer)
                            .hide();
                        return false;
                    } else {
                        checkForCheckout(props);
                    }

                    return false;
                });
        });
    };
    return {
        init: init
    }
}();