
var Visio;
(function (Visio) {
	var Application = (function(_super) {
		__extends(Application, _super);
		function Application() {
			/// <summary> Represents the Application. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="showBorders" type="Boolean">Show or hide the iFrame application borders. [Api set:  1.1]</field>
			/// <field name="showToolbars" type="Boolean">Show or hide the standard toolbars. [Api set:  1.1]</field>
		}

		Application.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Application"/>
		}

		Application.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.ApplicationUpdateData">Properties described by the Visio.Interfaces.ApplicationUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Application">An existing Application object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}
		Application.prototype.showToolbar = function(id, show) {
			/// <summary>
			/// Sets the visibility of a specific toolbar in the application. [Api set:  1.1]
			/// </summary>
			/// <param name="id" type="String">The type of the Toolbar</param>
			/// <param name="show" type="Boolean">Whether the toolbar is visibile or not.</param>
			/// <returns ></returns>
		}

		return Application;
	})(OfficeExtension.ClientObject);
	Visio.Application = Application;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var BoundingBox = (function() {
			function BoundingBox() {
				/// <summary> Represents the BoundingBox of the shape. [Api set:  1.1] </summary>
				/// <field name="height" type="Number">The distance between the top and bottom edges of the bounding box of the shape, excluding any data graphics associated with the shape. [Api set:  1.1]</field>
				/// <field name="width" type="Number">The distance between the left and right edges of the bounding box of the shape, excluding any data graphics associated with the shape. [Api set:  1.1]</field>
				/// <field name="x" type="Number">An integer that specifies the x-coordinate of the bounding box. [Api set:  1.1]</field>
				/// <field name="y" type="Number">An integer that specifies the y-coordinate of the bounding box. [Api set:  1.1]</field>
			}
			return BoundingBox;
		})();
		Interfaces.BoundingBox.__proto__ = null;
		Interfaces.BoundingBox = BoundingBox;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the type of column values. [Api set:  1.1] </summary>
	var ColumnType = {
		__proto__: null,
		"unknown": "unknown",
		"string": "string",
		"number": "number",
		"date": "date",
		"currency": "currency",
	}
	Visio.ColumnType = ColumnType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Comment = (function(_super) {
		__extends(Comment, _super);
		function Comment() {
			/// <summary> Represents the Comment. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="author" type="String">A string that specifies the name of the author of the comment. [Api set:  1.1]</field>
			/// <field name="date" type="String">A string that specifies the date when the comment was created. [Api set:  1.1]</field>
			/// <field name="text" type="String">A string that contains the comment text. [Api set:  1.1]</field>
		}

		Comment.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Comment"/>
		}

		Comment.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.CommentUpdateData">Properties described by the Visio.Interfaces.CommentUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOnReadOnly: Throw an error if the passed-in property list includes read-only properties (default = true).
			/// </param>
			/// </signature>
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on an existing loaded object.</summary>
			/// <param name="properties" type="Comment">An existing Comment object, with properties that have already been loaded and synced.</param>
			/// </signature>
		}

		return Comment;
	})(OfficeExtension.ClientObject);
	Visio.Comment = Comment;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var CommentCollection = (function(_super) {
		__extends(CommentCollection, _super);
		function CommentCollection() {
			/// <summary> Represents the CommentCollection for a given Shape. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="items" type="Array" elementType="Visio.Comment">Gets the loaded child items in this collection.</field>
		}

		CommentCollection.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.CommentCollection"/>
		}
		CommentCollection.prototype.getCount = function() {
			/// <summary>
			/// Gets the number of Comments. [Api set:  1.1]
			/// </summary>
			/// <returns type="OfficeExtension.ClientResult&lt;number&gt;"></returns>
			var result = new OfficeExtension.ClientResult();
			result.__proto__ = null;
			result.value = 0;
			return result;
		}
		CommentCollection.prototype.getItem = function(key) {
			/// <summary>
			/// Gets the Comment using its name. [Api set:  1.1]
			/// </summary>
			/// <param name="key" >Key is the name of the Comment to be retrieved.</param>
			/// <returns type="Visio.Comment"></returns>
		}

		return CommentCollection;
	})(OfficeExtension.ClientObject);
	Visio.CommentCollection = CommentCollection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var ConnectorBinding = (function() {
			function ConnectorBinding() {
				/// <summary> Connector bindings for data visualizer diagram. [Api set:  1.1] </summary>
				/// <field name="delimiter" type="String">Delimiter for TargetColumn. It should not have more then one character. [Api set:  1.1]</field>
			}
			return ConnectorBinding;
		})();
		Interfaces.ConnectorBinding.__proto__ = null;
		Interfaces.ConnectorBinding = ConnectorBinding;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Direction of connector in DataVisualizer diagram. [Api set:  1.1] </summary>
	var ConnectorDirection = {
		__proto__: null,
		"fromTarget": "fromTarget",
		"toTarget": "toTarget",
	}
	Visio.ConnectorDirection = ConnectorDirection;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the orientation of the Cross Functional Flowchart diagram. [Api set:  1.1] </summary>
	var CrossFunctionalFlowchartOrientation = {
		__proto__: null,
		"horizontal": "horizontal",
		"vertical": "vertical",
	}
	Visio.CrossFunctionalFlowchartOrientation = CrossFunctionalFlowchartOrientation;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Interfaces;
	(function (Interfaces) {
		var DataRefreshCompleteEventArgs = (function() {
			function DataRefreshCompleteEventArgs() {
				/// <summary> Provides information about the document that raised the DataRefreshComplete event. [Api set:  1.1] </summary>
				/// <field name="document" type="Visio.Document">Gets the document object that raised the DataRefreshComplete event. [Api set:  1.1]</field>
				/// <field name="success" type="Boolean">Gets the success or failure of the DataRefreshComplete event. [Api set:  1.1]</field>
			}
			return DataRefreshCompleteEventArgs;
		})();
		Interfaces.DataRefreshCompleteEventArgs.__proto__ = null;
		Interfaces.DataRefreshCompleteEventArgs = DataRefreshCompleteEventArgs;
	})(Interfaces = Visio.Interfaces || (Visio.Interfaces = { __proto__: null}));
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the type of source for the data connection. [Api set:  1.1] </summary>
	var DataSourceType = {
		__proto__: null,
		"unknown": "unknown",
		"excel": "excel",
	}
	Visio.DataSourceType = DataSourceType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Represents the types of data validation error. [Api set:  1.1] </summary>
	var DataValidationErrorType = {
		__proto__: null,
		"none": "none",
		"columnNotMapped": "columnNotMapped",
		"uniqueIdColumnError": "uniqueIdColumnError",
		"swimlaneColumnError": "swimlaneColumnError",
		"delimiterError": "delimiterError",
		"connectorColumnError": "connectorColumnError",
		"connectorColumnMappedElsewhere": "connectorColumnMappedElsewhere",
		"connectorLabelColumnMappedElsewhere": "connectorLabelColumnMappedElsewhere",
		"connectorColumnAndConnectorLabelMappedElsewhere": "connectorColumnAndConnectorLabelMappedElsewhere",
	}
	Visio.DataValidationErrorType = DataValidationErrorType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Type of the Data Visualizer Diagram operation [Api set:  1.1] </summary>
	var DataVisualizerDiagramOperationType = {
		__proto__: null,
		"unknown": "unknown",
		"create": "create",
		"updateMappings": "updateMappings",
		"updateData": "updateData",
		"update": "update",
		"delete": "delete",
	}
	Visio.DataVisualizerDiagramOperationType = DataVisualizerDiagramOperationType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> Result of Data Visualizer Diagram operations. [Api set:  1.1] </summary>
	var DataVisualizerDiagramResultType = {
		__proto__: null,
		"success": "success",
		"unexpected": "unexpected",
		"validationError": "validationError",
		"conflictError": "conflictError",
	}
	Visio.DataVisualizerDiagramResultType = DataVisualizerDiagramResultType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	/// <summary> DiagramType for Data Visualizer diagrams [Api set:  1.1] </summary>
	var DataVisualizerDiagramType = {
		__proto__: null,
		"unknown": "unknown",
		"basicFlowchart": "basicFlowchart",
		"crossFunctionalFlowchart_Horizontal": "crossFunctionalFlowchart_Horizontal",
		"crossFunctionalFlowchart_Vertical": "crossFunctionalFlowchart_Vertical",
		"audit": "audit",
		"orgChart": "orgChart",
		"network": "network",
	}
	Visio.DataVisualizerDiagramType = DataVisualizerDiagramType;
})(Visio || (Visio = {__proto__: null}));

var Visio;
(function (Visio) {
	var Document = (function(_super) {
		__extends(Document, _super);
		function Document() {
			/// <summary> Represents the Document class. [Api set:  1.1] </summary>
			/// <field name="context" type="Visio.RequestContext">The request context associated with this object.</field>
			/// <field name="isNull" type="Boolean">Returns a boolean value for whether the corresponding object is null. You must call "context.sync()" before reading the isNull property.</field>
			/// <field name="application" type="Visio.Application">Represents a Visio application instance that contains this document. Read-only. [Api set:  1.1]</field>
			/// <field name="pages" type="Visio.PageCollection">Represents a collection of pages associated with the document. Read-only. [Api set:  1.1]</field>
			/// <field name="view" type="Visio.DocumentView">Returns the DocumentView object. Read-only. [Api set:  1.1]</field>
			/// <field name="onDataRefreshComplete" type="OfficeExtension.EventHandlers">Occurs when the data is refreshed in the diagram. [Api set:  1.1]</field>
			/// <field name="onDocumentError" type="OfficeExtension.EventHandlers">Occurs when there is an expected or unexpected error occured in the session. [Api set:  1.1]</field>
			/// <field name="onDocumentLoadComplete" type="OfficeExtension.EventHandlers">Occurs when the Document is loaded, refreshed, or changed. [Api set:  1.1]</field>
			/// <field name="onPageLoadComplete" type="OfficeExtension.EventHandlers">Occurs when the page is finished loading. [Api set:  1.1]</field>
			/// <field name="onSelectionChanged" type="OfficeExtension.EventHandlers">Occurs when the current selection of shapes changes. [Api set:  1.1]</field>
			/// <field name="onShapeMouseEnter" type="OfficeExtension.EventHandlers">Occurs when the user moves the mouse pointer into the bounding box of a shape. [Api set:  1.1]</field>
			/// <field name="onShapeMouseLeave" type="OfficeExtension.EventHandlers">Occurs when the user moves the mouse out of the bounding box of a shape. [Api set:  1.1]</field>
			/// <field name="onTaskPaneStateChanged" type="OfficeExtension.EventHandlers">Occurs whenever a task pane state is changed [Api set:  1.1]</field>
		}

		Document.prototype.load = function(option) {
			/// <summary>
			/// Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
			/// </summary>
			/// <param name="option" type="string | string[] | OfficeExtension.LoadOption"/>
			/// <returns type="Visio.Document"/>
		}

		Document.prototype.set = function() {
			/// <signature>
			/// <summary>Sets multiple properties on the object at the same time, based on JSON input.</summary>
			/// <param name="properties" type="Visio.Interfaces.DocumentUpdateData">Properties described by the Visio.Interfaces.DocumentUpdateData interface.</param>
			/// <param name="options" type="string">Options of the form { throwOnReadOnly?: boolean }
			/// <br />
			/// * throwOn