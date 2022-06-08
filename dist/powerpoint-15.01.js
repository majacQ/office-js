/* PowerPoint specific API library */
/* Version: 15.0.4764.1000 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

OSF.ClientMode={ReadWrite:0,ReadOnly:1};OSF.DDA.RichInitializationReason={1:Microsoft.Office.WebExtension.InitializationReason.Inserted,2:Microsoft.Office.WebExtension.InitializationReason.DocumentOpened};Microsoft.Office.WebExtension.FileType={Text:"text",Compressed:"compressed"};OSF.DDA.RichClientSettingsManager={read:function(e,d){var b=[],f=[];e&&e();window.external.GetContext().GetSettings().Read(b,f);d&&d();for(var c={},a=0;a<b.length;a++)c[b[a]]=f[a];return c},write:function(a,g,c,b){var e=[],d=[];for(var f in a){e.push(f);d.push(a[f])}c&&c();window.external.GetContext().GetSettings().Write(e,d);b&&b()}};OSF.DDA.DispIdHost.getRichClientDelegateMethods=function(e){var a={};a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=OSF.DDA.SafeArray.Delegate.executeAsync;a[OSF.DDA.DispIdHost.Delegates.RegisterEventAsync]=OSF.DDA.SafeArray.Delegate.registerEventAsync;a[OSF.DDA.DispIdHost.Delegates.UnregisterEventAsync]=OSF.DDA.SafeArray.Delegate.unregisterEventAsync;function b(a){return function(b){var d,c;try{c=a(b.hostCallArgs,b.onCalling,b.onReceiving);d=OSF.DDA.ErrorCodeManager.errorCodes.ooeSuccess}catch(e){d=OSF.DDA.ErrorCodeManager.errorCodes.ooeInternalError;c={name:Strings.OfficeOM.L_InternalError,message:e}}b.onComplete&&b.onComplete(d,c)}}function d(c,b,a){return OSF.DDA.RichClientSettingsManager.read(b,a)}function c(a,c,b){return OSF.DDA.RichClientSettingsManager.write(a[OSF.DDA.SettingsManager.SerializedSettings],a[Microsoft.Office.WebExtension.Parameters.OverwriteIfStale],c,b)}switch(e){case OSF.DDA.AsyncMethodNames.RefreshAsync.id:a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=b(d);break;case OSF.DDA.AsyncMethodNames.SaveAsync.id:a[OSF.DDA.DispIdHost.Delegates.ExecuteAsync]=b(c)}return a};OSF.DDA.File=function(e,c,b){OSF.OUtil.defineEnumerableProperties(this,{size:{value:c},sliceCount:{value:Math.ceil(c/b)}});var a={};a[OSF.DDA.FileProperties.Handle]=e;a[OSF.DDA.FileProperties.SliceSize]=b;var d=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[d.GetDocumentCopyChunkAsync,d.ReleaseDocumentCopyAsync],a)};OSF.DDA.FileSliceOffset="fileSliceoffset";OSF.DDA.CustomXmlParts=function(){this._eventDispatches=[];var a=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[a.AddDataPartAsync,a.GetDataPartByIdAsync,a.GetDataPartsByNameSpaceAsync])};OSF.DDA.CustomXmlPart=function(f,b,g){OSF.OUtil.defineEnumerableProperties(this,{builtIn:{value:g},id:{value:b},namespaceManager:{value:new OSF.DDA.CustomXmlPrefixMappings(b)}});var c=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[c.DeleteDataPartAsync,c.GetPartNodesAsync,c.GetPartXmlAsync]);var e=f._eventDispatches,a=e[b];if(!a){var d=Microsoft.Office.WebExtension.EventType;a=new OSF.EventDispatch([d.DataNodeDeleted,d.DataNodeInserted,d.DataNodeReplaced]);e[b]=a}OSF.DDA.DispIdHost.addEventSupport(this,a)};OSF.DDA.CustomXmlPrefixMappings=function(b){var a=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[a.AddDataPartNamespaceAsync,a.GetDataPartNamespaceAsync,a.GetDataPartPrefixAsync],b)};OSF.DDA.CustomXmlNode=function(d,c,e,b){OSF.OUtil.defineEnumerableProperties(this,{baseName:{value:b},namespaceUri:{value:e},nodeType:{value:c}});var a=OSF.DDA.AsyncMethodNames;OSF.DDA.DispIdHost.addAsyncMethods(this,[a.GetRelativeNodesAsync,a.GetNodeValueAsync,a.GetNodeXmlAsync,a.SetNodeValueAsync,a.SetNodeXmlAsync],d)};OSF.DDA.NodeInsertedEventArgs=function(b,a){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DataNodeInserted},newNode:{value:b},inUndoRedo:{value:a}})};OSF.DDA.NodeReplacedEventArgs=function(c,b,a){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DataNodeReplaced},oldNode:{value:c},newNode:{value:b},inUndoRedo:{value:a}})};OSF.DDA.NodeDeletedEventArgs=function(c,a,b){OSF.OUtil.defineEnumerableProperties(this,{type:{value:Microsoft.Office.WebExtension.EventType.DataNodeDeleted},oldNode:{value:c},oldNextSibling:{value:a},inUndoRedo:{value:b}})};OSF.OUtil.redefineList(Microsoft.Office.WebExtension.FileType,{Compressed:"compressed",Pdf:"pdf"});OSF.OUtil.redefineList(Microsoft.Office.WebExtension.CoercionType,{Text:"text",SlideRange:"slideRange"});OSF.OUtil.redefineList(Microsoft.Office.WebExtension.EventType,{DocumentSelectionChanged:"documentSelectionChanged",OfficeThemeChanged:"officeThemeChanged",DocumentThemeChanged:"documentThemeChanged"});OSF.OUtil.redefineList(Microsoft.Office.WebExtension.EventType,{DocumentSelectionChanged:"documentSelectionChanged",OfficeThemeChanged:"officeThemeChanged",DocumentThemeChanged:"documentThemeChanged",ActiveViewChanged:"activeViewChanged"});OSF.OUtil.redefineList(Microsoft.Office.WebExtension.ValueFormat,{Unformatted:"unformatted"});OSF.OUtil.redefineList(Microsoft.Office.WebExtension.FilterType,{All:"all"});Microsoft.Office.Internal.OfficeTheme={PrimaryFontColor:"primaryFontColor",PrimaryBackgroundColor:"primaryBackgroundColor",SecondaryFontColor:"secondaryFontColor",SecondaryBackgroundColor:"secondaryBackgroundColor"};Microsoft.Office.Internal.DocumentTheme={PrimaryFontColor:"primaryFontColor",PrimaryBackgroundColor:"primaryBackgroundColor",SecondaryFontColor:"secondaryFontColor",SecondaryBackgroundColor:"secondaryBackgroundColor",Accent1:"accent1",Accent2:"accent2",Accent3:"accent3",Accent4:"accent4",Accent5:"accent5",Accent6:"accent6",Hyperlink:"hyperlink",FollowedHyperlink:"followedHyperlink",HeaderLatinFont:"headerLatinFont",HeaderEastAsianFont:"headerEastAsianFont",HeaderScriptFont:"headerScriptFont",HeaderLocalizedFont:"headerLocalizedFont",BodyLatinFont:"bodyLatinFont",BodyEastAsianFont:"bodyEastAsianFont",BodyScriptFont:"bodyScriptFont",BodyLocalizedFont:"bodyLocalizedFont"};Microsoft.Office.WebExtension.ActiveView={};OSF.OUtil.redefineList(Microsoft.Office.WebExtension.ActiveView,{Read:"read",Edit:"edit"});OSF.OUtil.redefineList(Microsoft.Office.WebExtension.GoToType,{Slide:"slide",Index:"index"});delete Microsoft.Office.WebExtension.BindingType;delete Microsoft.Office.WebExtension.select;OSF.OUtil.setNamespace("SafeArray",OSF.DDA);OSF.DDA.SafeArray.Response={Status:0,Payload:1};OSF.DDA.SafeArray.UniqueArguments={Offset:"offset",Run:"run",BindingSpecificData:"bindingSpecificData",MergedCellGuid:"{66e7831f-81b2-42e2-823c-89e872d541b3}"};OSF.OUtil.setNamespace("Delegate",OSF.DDA.SafeArray);OSF.DDA.SafeArray.Delegate.SpecialProcessor=function(){var b=this;function c(a){var b;try{var h=a.ubound(1),d=a.ubound(2);a=a.toArray();if(h==1&&d==1)b=[a];else{b=[];for(var f=0;f<h;f++){for(var c=[],e=0;e<d;e++){var g=a[f*d+e];g!=OSF.DDA.SafeArray.UniqueArguments.MergedCellGuid&&c.push(g)}c.length>0&&b.push(c)}}}catch(i){}return b}var d=[OSF.DDA.PropertyDescriptors.FileProperties,OSF.DDA.PropertyDescriptors.FileSliceProperties,OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,OSF.DDA.PropertyDescriptors.BindingProperties,OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,OSF.DDA.SafeArray.UniqueArguments.Offset,OSF.DDA.SafeArray.UniqueArguments.Run,OSF.DDA.PropertyDescriptors.Subset,OSF.DDA.PropertyDescriptors.DataPartProperties,OSF.DDA.PropertyDescriptors.DataNodeProperties,OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,OSF.DDA.EventDescriptors.DataNodeInsertedEvent,OSF.DDA.EventDescriptors.DataNodeReplacedEvent,OSF.DDA.EventDescriptors.DataNodeDeletedEvent,OSF.DDA.EventDescriptors.DocumentThemeChangedEvent,OSF.DDA.EventDescriptors.OfficeThemeChangedEvent,OSF.DDA.EventDescriptors.ActiveViewChangedEvent,OSF.DDA.DataNodeEventProperties.OldNode,OSF.DDA.DataNodeEventProperties.NewNode,OSF.DDA.DataNodeEventProperties.NextSiblingNode,Microsoft.Office.Internal.Parameters.OfficeTheme,Microsoft.Office.Internal.Parameters.DocumentTheme],a={};a[Microsoft.Office.WebExtension.Parameters.Data]=function(){var b=0,a=1;return {toHost:function(c){if(typeof c!="string"&&c[OSF.DDA.TableDataProperties.TableRows]!==undefined){var d=[];d[b]=c[OSF.DDA.TableDataProperties.TableRows];d[a]=c[OSF.DDA.TableDataProperties.TableHeaders];c=d}return c},fromHost:function(f){var e;if(f.toArray){var g=f.dimensions();if(g===2)e=c(f);else{var d=f.toArray();if(d.length===2&&(d[0]!=null&&d[0].toArray||d[1]!=null&&d[1].toArray)){e={};e[OSF.DDA.TableDataProperties.TableRows]=c(d[b]);e[OSF.DDA.TableDataProperties.TableHeaders]=c(d[a])}else e=d}}else e=f;return e}}}();OSF.DDA.SafeArray.Delegate.SpecialProcessor.uber.constructor.call(b,d,a);b.pack=function(c,d){var b;if(this.isDynamicType(c))b=a[c].toHost(d);else b=d;return b};b.unpack=function(c,d){var b;if(this.isComplexType(c)||OSF.DDA.ListType.isListType(c))try{b=d.toArray()}catch(e){b=d||{}}else if(this.isDynamicType(c))b=a[c].fromHost(d);else b=d;return b};b.dynamicTypes=a};OSF.OUtil.extend(OSF.DDA.SafeArray.Delegate.SpecialProcessor,OSF.DDA.SpecialProcessor);OSF.DDA.SafeArray.Delegate.ParameterMap=function(){var f=true,e=new OSF.DDA.HostParameterMap(new OSF.DDA.SafeArray.Delegate.SpecialProcessor),a,d=e.self;function g(a){var c=null;if(a){c={};for(var d=a.length,b=0;b<d;b++)c[a[b].name]=a[b].value}return c}function b(b){var a={},c=g(b.toHost);if(b.invertible)a.map=c;else if(b.canonical)a.toHost=a.fromHost=c;else{a.toHost=c;a.fromHost=g(b.fromHost)}e.setMapping(b.type,a)}a=OSF.DDA.FileProperties;b({type:OSF.DDA.PropertyDescriptors.FileProperties,fromHost:[{name:a.Handle,value:0},{name:a.FileSize,value:1}]});b({type:OSF.DDA.PropertyDescriptors.FileSliceProperties,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.Data,value:0},{name:a.SliceSize,value:1}]});a=OSF.DDA.FilePropertiesDescriptor;b({type:OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,fromHost:[{name:a.Url,value:0}]});a=OSF.DDA.BindingProperties;b({type:OSF.DDA.PropertyDescriptors.BindingProperties,fromHost:[{name:a.Id,value:0},{name:a.Type,value:1},{name:OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,value:2}]});b({type:OSF.DDA.SafeArray.UniqueArguments.BindingSpecificData,fromHost:[{name:a.RowCount,value:0},{name:a.ColumnCount,value:1},{name:a.HasHeaders,value:2}]});a=OSF.DDA.SafeArray.UniqueArguments;b({type:OSF.DDA.PropertyDescriptors.Subset,toHost:[{name:a.Offset,value:0},{name:a.Run,value:1}],canonical:f});a=Microsoft.Office.WebExtension.Parameters;b({type:OSF.DDA.SafeArray.UniqueArguments.Offset,toHost:[{name:a.StartRow,value:0},{name:a.StartColumn,value:1}],canonical:f});b({type:OSF.DDA.SafeArray.UniqueArguments.Run,toHost:[{name:a.RowCount,value:0},{name:a.ColumnCount,value:1}],canonical:f});a=OSF.DDA.DataPartProperties;b({type:OSF.DDA.PropertyDescriptors.DataPartProperties,fromHost:[{name:a.Id,value:0},{name:a.BuiltIn,value:1}]});a=OSF.DDA.DataNodeProperties;b({type:OSF.DDA.PropertyDescriptors.DataNodeProperties,fromHost:[{name:a.Handle,value:0},{name:a.BaseName,value:1},{name:a.NamespaceUri,value:2},{name:a.NodeType,value:3}]});b({type:OSF.DDA.EventDescriptors.BindingSelectionChangedEvent,fromHost:[{name:OSF.DDA.PropertyDescriptors.BindingProperties,value:0},{name:OSF.DDA.PropertyDescriptors.Subset,value:1}]});b({type:OSF.DDA.EventDescriptors.DocumentThemeChangedEvent,fromHost:[{name:Microsoft.Office.Internal.Parameters.DocumentTheme,value:d}]});b({type:OSF.DDA.EventDescriptors.OfficeThemeChangedEvent,fromHost:[{name:Microsoft.Office.Internal.Parameters.OfficeTheme,value:d}]});b({type:OSF.DDA.EventDescriptors.ActiveViewChangedEvent,fromHost:[{name:Microsoft.Office.WebExtension.Parameters.ActiveView,value:0}]});a=OSF.DDA.DataNodeEventProperties;b({type:OSF.DDA.EventDescriptors.DataNodeInsertedEvent,fromHost:[{name:a.InUndoRedo,value:0},{name:a.NewNode,value:1}]});b({type:OSF.DDA.EventDescriptors.DataNodeReplacedEvent,fromHost:[{name:a.InUndoRedo,value:0},{name:a.OldNode,value:1},{name:a.NewNode,value:2}]});b({type:OSF.DDA.EventDescriptors.DataNodeDeletedEvent,fromHost:[{name:a.InUndoRedo,value:0},{name:a.OldNode,value:1},{name:a.NextSiblingNode,value:2}]});b({type:a.OldNode,fromHost:[{name:OSF.DDA.PropertyDescriptors.DataNodeProperties,value:d}]});b({type:a.NewNode,fromHost:[{name:OSF.DDA.PropertyDescriptors.DataNodeProperties,value:d}]});b({type:a.NextSiblingNode,fromHost:[{name:OSF.DDA.PropertyDescriptors.DataNodeProperties,value:d}]});a=Microsoft.Office.WebExtension.AsyncResultStatus;b({type:OSF.DDA.PropertyDescriptors.AsyncResultStatus,fromHost:[{name:a.Succeeded,value:0},{name:a.Failed,value:1}]});a=Microsoft.Office.WebExtension.CoercionType;b({type:Microsoft.Office.WebExtension.Parameters.CoercionType,toHost:[{name:a.Text,value:0},{name:a.Matrix,value:1},{name:a.Table,value:2},{name:a.Html,value:3},{name:a.Ooxml,value:4},{name:a.SlideRange,value:7}]});a=Microsoft.Office.WebExtension.GoToType;b({type:Microsoft.Office.WebExtension.Parameters.GoToType,toHost:[{name:a.Binding,value:0},{name:a.NamedItem,value:1},{name:a.Slide,value:2},{name:a.Index,value:3}]});a=Microsoft.Office.WebExtension.FileType;a&&b({type:Microsoft.Office.WebExtension.Parameters.FileType,toHost:[{name:a.Text,value:0},{name:a.Compressed,value:5},{name:a.Pdf,value:6}]});a=Microsoft.Office.WebExtension.BindingType;a&&b({type:Microsoft.Office.WebExtension.Parameters.BindingType,toHost:[{name:a.Text,value:0},{name:a.Matrix,value:1},{name:a.Table,value:2}],invertible:f});a=Microsoft.Office.WebExtension.ValueFormat;b({type:Microsoft.Office.WebExtension.Parameters.ValueFormat,toHost:[{name:a.Unformatted,value:0},{name:a.Formatted,value:1}]});a=Microsoft.Office.WebExtension.FilterType;b({type:Microsoft.Office.WebExtension.Parameters.FilterType,toHost:[{name:a.All,value:0},{name:a.OnlyVisible,value:1}]});a=Microsoft.Office.Internal.OfficeTheme;a&&b({type:Microsoft.Office.Internal.Parameters.OfficeTheme,fromHost:[{name:a.PrimaryFontColor,value:0},{name:a.PrimaryBackgroundColor,value:1},{name:a.SecondaryFontColor,value:2},{name:a.SecondaryBackgroundColor,value:3}]});a=Microsoft.Office.WebExtension.ActiveView;a&&b({type:Microsoft.Office.WebExtension.Parameters.ActiveView,fromHost:[{name:0,value:a.Read},{name:1,value:a.Edit}]});a=Microsoft.Office.Internal.DocumentTheme;a&&b({type:Microsoft.Office.Internal.Parameters.DocumentTheme,fromHost:[{name:a.PrimaryBackgroundColor,value:0},{name:a.PrimaryFontColor,value:1},{name:a.SecondaryBackgroundColor,value:2},{name:a.SecondaryFontColor,value:3},{name:a.Accent1,value:4},{name:a.Accent2,value:5},{name:a.Accent3,value:6},{name:a.Accent4,value:7},{name:a.Accent5,value:8},{name:a.Accent6,value:9},{name:a.Hyperlink,value:10},{name:a.FollowedHyperlink,value:11},{name:a.HeaderLatinFont,value:12},{name:a.HeaderEastAsianFont,value:13},{name:a.HeaderScriptFont,value:14},{name:a.HeaderLocalizedFont,value:15},{name:a.BodyLatinFont,value:16},{name:a.BodyEastAsianFont,value:17},{name:a.BodyScriptFont,value:18},{name:a.BodyLocalizedFont,value:19}]});a=Microsoft.Office.WebExtension.SelectionMode;b({type:Microsoft.Office.WebExtension.Parameters.SelectionMode,toHost:[{name:a.Default,value:0},{name:a.Selected,value:1},{name:a.None,value:2}]});a=Microsoft.Office.WebExtension.Parameters;var c=OSF.DDA.MethodDispId;b({type:c.dispidNavigateToMethod,toHost:[{name:a.Id,value:0},{name:a.GoToType,value:1},{name:a.SelectionMode,value:2}]});b({type:c.dispidGetSelectedDataMethod,fromHost:[{name:a.Data,value:d}],toHost:[{name:a.CoercionType,value:0},{name:a.ValueFormat,value:1},{name:a.FilterType,value:2}]});b({type:c.dispidSetSelectedDataMethod,toHost:[{name:a.CoercionType,value:0},{name:a.Data,value:1}]});b({type:c.dispidGetFilePropertiesMethod,fromHost:[{name:OSF.DDA.PropertyDescriptors.FilePropertiesDescriptor,value:d}]});b({type:c.dispidGetDocumentCopyMethod,toHost:[{name:a.FileType,value:0}],fromHost:[{name:OSF.DDA.PropertyDescriptors.FileProperties,value:d}]});b({type:c.dispidGetDocumentCopyChunkMethod,toHost:[{name:OSF.DDA.FileProperties.Handle,value:0},{name:OSF.DDA.FileSliceOffset,value:1},{name:OSF.DDA.FileProperties.SliceSize,value:2}],fromHost: