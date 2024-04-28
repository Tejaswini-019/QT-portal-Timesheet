export enum ListTemplateType {
    CustomList = 100, //A basic list that can be adapted for multiple purposes.
    DocumentLibrary = 101, //Contains a list of documents and other files.
    Survey = 102, //Fields (2) on a survey list represent questions that are asked of survey participants. Items in a list represent a set of responses to a survey.
    Links = 103, //Contains a list of hyperlinks and their descriptions.
    Announcements = 104, //Contains a set of simple announcements.
    Contacts = 105, //Contains a list of contacts used for tracking people in a site (2).
    Calendar = 106, //Contains a list of single and recurring events. An events list has special views for displaying events on a calendar.
    Tasks = 107, //Contains a list of items that represent finished and pending work items.
    DiscussionBoard = 108, //Contains discussions entries and their replies.
    PictureLibrary = 109, //Contains a library adapted for storing and viewing digital pictures.
    DataSources = 110, //Contains data connection description files.
    FormLibrary = 115, //Contains XML documents. An XML form library can also contain templates for displaying and editing XML files through forms, as well as rules for specifying how XML data is converted to and from list items.
    NoCodeWorkflows = 117, //Contains additional workflow definitions that describe new processes that can be used in lists. These workflow definitions do not contain advanced code-based extensions.
    CustomWorkflowProcess = 118, //Contains a list used to support custom workflow process actions.
    WikiPageLibrary = 119, //Contains a set of editable Web pages.
    CustomGrid = 120, //Contains a set of list items with a grid-editing view.
    NoCodePublicWorkflows = 122, //A gallery for storing workflow definitions that do not contain advanced code-based extensions.
    WorkflowHistory = 140, //Contains a set of history items for instances of workflows.
    ProjectTasks = 150, //Contains a list of tasks with specialized views of task data in the form of Gantt chart.
    PublicWorkflowsExternalList = 600, //An external list for viewing the data of an external content type.
    IssuesTracking = 1100, //Contains a list of items to track issues.
}

export enum TemplateFileType {
    StandardPage = 0,
    WikiPage = 1,
    FormPage = 2,
    ClientSidePage = 3
}

export enum FieldTypeKind {
    Invalid = 0, //Must not be used. The value = 0.
    Integer = 1, //Specifies that the field contains an integer value.
    Text = 2, //Specifies that the field contains a single line of text. SP.FieldText
    Note = 3, //Specifies that the field contains multiple lines of text.
    DateTime = 4, //Specifies that the field contains a date and time value or a date-only value. SP.FieldDateTime
    Counter = 5, //Specifies that the field contains a monotonically increasing integer.
    Choice = 6, //Specifies that the field contains a single value from a set of specified values.
    Lookup = 7, //Specifies that the field is a lookup field.
    Boolean = 8, //Specifies that the field contains a Boolean value.
    Number = 9, //Specifies that the field contains a floating-point number value. SP.FieldNumber
    Currency = 10, //Specifies that the field contains a currency value. SP.FieldCurrency
    URL = 11, //Specifies that the field contains a URI and an optional description of the URI.
    Computed = 12, //Specifies that the field is a computed field.
    Threading = 13, //Specifies that the field indicates the thread for a discussion item in a threaded view of a discussion board.
    Guid = 14, //Specifies that the field contains a GUID value.
    MultiChoice = 15, //Specifies that the field contains one or more values from a set of specified values.
    GridChoice = 16, //Specifies that the field contains rating scale values for a survey list. The value = 16.
    Calculated = 17, //Specifies that the field is a calculated field. The value = 17.
    File = 18, //Specifies that the field contains the leaf name of a document as a value. The value = 18.
    Attachments = 19,//Specifies that the field indicates whether the list item has attachments. The value = 19.
    User = 20, //Specifies that the field contains one or more users and groups as values. The value = 20.
    Recurrence = 21, //Specifies that the field indicates whether a meeting in a calendar list recurs. The value = 21.
    CrossProjectLink = 22, //Specifies that the field contains a link between projects in a Meeting Workspace site. The value = 22.
    ModStat = 23, //Specifies that the field indicates moderation status. The value = 23.
    Error = 24, //Specifies that the type of the field was set to an invalid value. The value = 24.
    ContentTypeId = 25, //Specifies that the field contains a content type identifier as a value. The value = 25.
    PageSeparator = 26, //Specifies that the field separates questions in a survey list onto multiple pages. The value = 26.
    ThreadIndex = 27,//Specifies that the field indicates the position of a discussion item in a threaded view of a discussion board. The value = 27.
    WorkflowStatus = 28,//Specifies that the field indicates the status of a workflow instance on a list item. The value = 28.
    AllDayEvent = 29, //Specifies that the field indicates whether a meeting in a calendar list is an all-day event. The value = 29.
    WorkflowEventType = 30, //Specifies that the field contains the most recent event in a workflow instance. The value = 30.
    MaxItems = 31 //Must not be used. The value = 31.
}

export enum CalendarEventPeriod {
    Today,
    Month,
    Year
}

export enum Status {
    Draft = "Draft",
    Approver1 = "Pending with Approver 1",
    Approver2 = "Pending with Approver 2",
    Approver3 = "Pending with Approver 3",
    Approver4 = "Pending with Approver 4",
    Approver5 = "Pending with Approver 5",
    Rejected = "Rejected",
    Approved = "Approved",
    Cancelled = "Cancelled",
    Resubmission = "Pending for Resubmission"
}


export enum Action {
    Submitted = "Submitted",
    Approved = "Approved",
    RequestedforAmendment = "Requested for Amendment",
    Delegated = "Delegated",
    Approver4 = "Pending with Approver 4",
    Approver5 = "Pending with Approver 5",
    Rejected = "Rejected",
    Cancelled = "Cancelled",
    Resubmitted = "Resubmitted",
    Delegation = "Delegation"
}

export enum UrlTabs {
    Draft = "Draft",
    MyRequests = "MyRequests",
    Resubmissions = "Resubmissions",
    ActionedRequests = "ActionedRequests",
    AllRequests = "AllRequests",
}