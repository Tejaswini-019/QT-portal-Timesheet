import * as React from 'react';
import './MyTimesheet.module.scss';
import { IMyTimesheetProps } from './IMyTimesheetProps';
import SPService from '../../../shared/services/SPService';
import * as moment from 'moment';
import * as $ from 'jquery';
import { applyDataTable, destroyDataTable, applyDataTableGroupBy } from '../../../shared/utils/datatable';
import { download, downloadWithGroupBy } from '../../../shared/utils/datatable-export';
import { IValidationField } from '../../../shared/models/IValidationField';
import { filter, orderBy } from 'lodash';
import "select2/dist/css/select2.min.css";
import "select2/dist/js/select2.min.js";
import 'jquery-ui-dist/jquery-ui.min.js';
import 'jquery-ui-dist/jquery-ui.min.css';
require("../../../shared/js/summernote-bs4.js");
import { initializeIcons } from '@fluentui/react/lib/Icons';
initializeIcons();
import FullCalendar, { formatDate } from '@fullcalendar/react' // must go before plugins
import dayGridPlugin from '@fullcalendar/daygrid' // a plugin!
import timeGridPlugin from '@fullcalendar/timegrid'
import interactionPlugin from '@fullcalendar/interaction'
import { INITIAL_EVENTS, createEventId } from './event-utils'
import { ColorPicker } from 'office-ui-fabric-react';
import { MSGraphClient } from '@microsoft/sp-http';
let tooltipInstance = null;
export interface IMyTimesheetState {
  EmployeeList: any[];
  ProjectList: any[];
  project: any;
  dateRange: any;
  fromDate: any;
  toDate: any;
  errorMessageFromDate: string;
  errorMessageToDate: string;
  TaskType: any[];
  status: string;
  statusFilter: string;
  successMessage: string;
  alertMessage: string;
  screen: boolean;
  export: boolean;
  isPending: boolean;
  isCalender: boolean;
  efforts: any;
  isManager: boolean;
  createdBy: any;
  groupColumn: any;
  weekendsVisible: true,
  currentEvents: []
  event: any,
  ManagerEMail: any;
  ManagerTitle: any;
}

export default class MyTimesheet extends React.Component<IMyTimesheetProps, IMyTimesheetState> {
  private _spservice: SPService;
  constructor(props) {
    super(props);
    this._spservice = new SPService(this.props.context);
    this.state = {
      EmployeeList: [],
      ProjectList: [],
      project: [],
      dateRange: 0,
      fromDate: "",
      toDate: "",
      errorMessageFromDate: "",
      errorMessageToDate: "",
      TaskType: [],
      status: "",
      statusFilter: "",
      successMessage: "",
      alertMessage: "",
      screen: false,
      export: false,
      isPending: true,
      efforts: "",
      isCalender: false,
      isManager: false,
      createdBy: "",
      groupColumn: "",
      weekendsVisible: true,
      currentEvents: [],
      event: [],
      ManagerEMail: "",
      ManagerTitle: ""
    }
  }

  public async componentDidMount() {
    $("#calendar").hide();
    this.toggleLoader(true);
    let _this = this;
    await this.getProjects();
    await this.getTaskType();
    if ($('#pending').hasClass('active')) {
      await this.getPendingTimesheets();
    } else if ($('#all').hasClass('active')) {
      await this.getAllTimesheets();
    }
    $(document).on('change', '#projectName2', function () {
      let id = $('#projectName2').val();
      _this.handleSelectedId(id);
    });
    $(document).on('change', '#projectName3', function () {
      let id = $('#projectName3').val();
      _this.handleSelectedId(id);
    });
    ($('#description') as any).summernote({
      shortcuts: false,
      addDefaultFonts: true,
      height: 200,
      toolbar: [
        ['para', ['ul', 'ol']],
      ],
      callbacks: {
        onPaste: function (e) {
          var bufferText = ((e.originalEvent || e).clipboardData).getData('text/html') || ((e.originalEvent || e).clipboardData).getData('text/plain');
          e.preventDefault();
          var div = $('<div />');
          div.append(bufferText);
          div.find('*').removeAttr('style');
          setTimeout(function () {
            document.execCommand('insertHtml', false, div.html());
          }, 10);
        }
      }
    });
    ($('#description1') as any).summernote({
      shortcuts: false,
      addDefaultFonts: true,
      height: 200,
      toolbar: [
        ['para', ['ul', 'ol']],
      ],
      callbacks: {
        onPaste: function (e) {
          var bufferText = ((e.originalEvent || e).clipboardData).getData('text/html') || ((e.originalEvent || e).clipboardData).getData('text/plain');
          e.preventDefault();
          var div = $('<div />');
          div.append(bufferText);
          div.find('*').removeAttr('style');
          setTimeout(function () {
            document.execCommand('insertHtml', false, div.html());
          }, 10);
        }
      }
    });
    $(document).on('change', 'input#selectprojectname', function (e) {
      e.preventDefault();
      var val = $("#selectprojectname").val();
      var projectID = $('#projectsandclients option').filter(function () {
        return this.value == val;
      }).data('id');
      if (projectID) {
        $("#projectName").val(projectID);
      } else {
        $("#projectName").val("");
        $("#selectprojectname").val("");
      }
    });
    $(document).on('change', 'input#selectprojectname1', function (e) {
      e.preventDefault();
      var val = $("#selectprojectname1").val();
      var projectID = $('#projectsandclients1 option').filter(function () {
        return this.value == val;
      }).data('id');
      if (projectID) {
        $("#projectName1").val(projectID);
      } else {
        $("#projectName1").val("");
        $("#selectprojectname1").val("");
      }
    });
  }
  private async loadDefaultDates(option) {
    let opt = option;
    //let opt = $('#daterange').val();
    //let opt = this.state.dateRange;
    const now = moment();
    console.log(now, "date")
    switch (opt) {
      case "1":
        //Today
        let today = now.format("YYYY-MM-DD");
        this.setState({
          fromDate: today ? today : "",
          toDate: today ? today : ""
        });
        break;
      case "0":
        //this month
        let thisMonthsFromDate = moment(now).startOf('month').format('YYYY-MM-DD');
        // let thisMonthsToDate = moment(now).subtract(1, 'months').endOf('month').format('YYYY-MM-DD');
        let thisMonthsToDate = now.format("YYYY-MM-DD");
        this.setState({
          fromDate: thisMonthsFromDate ? thisMonthsFromDate : "",
          toDate: thisMonthsToDate ? thisMonthsToDate : ""
        });
        break;
      case "2":
        //last month					 
        let lastMonthFromDate = moment(now).subtract(1, 'months').startOf('month').format('YYYY-MM-DD');
        let lastMonthToDate = moment(now).subtract(1, 'months').endOf('month').format('YYYY-MM-DD');
        // let lastMonthToDate = now.format("YYYY-MM-DD");
        this.setState({
          fromDate: lastMonthFromDate ? lastMonthFromDate : "",
          toDate: lastMonthToDate ? lastMonthToDate : ""
        });
        break;
      case "4":
        //last two months
        let lasttwoMonthsFromDate = moment(now).subtract(2, 'months').startOf('month').format('YYYY-MM-DD');
        let lasttwoMonthsToDate = moment(now).subtract(1, 'months').endOf('month').format('YYYY-MM-DD');
        // let lasttwoMonthsToDate = now.format("YYYY-MM-DD");
        this.setState({
          fromDate: lasttwoMonthsFromDate ? lasttwoMonthsFromDate : "",
          toDate: lasttwoMonthsToDate ? lasttwoMonthsToDate : ""
        });
        break;
      case "5":
        //last three months
        let lastthreeMonthsFromDate = moment(now).subtract(3, 'months').startOf('month').format('YYYY-MM-DD');
        let lastthreeMonthsToDate = moment(now).subtract(1, 'months').endOf('month').format('YYYY-MM-DD');
        // let lastthreeMonthsToDate = now.format("YYYY-MM-DD");
        this.setState({
          fromDate: lastthreeMonthsFromDate ? lastthreeMonthsFromDate : "",
          toDate: lastthreeMonthsToDate ? lastthreeMonthsToDate : ""
        });
        break;
      // case "3":			
      //   //custom
      //   this.setState({ 
      //     fromDate: "",
      //     toDate: ""
      //   });
      // break;      
    }
  }
  private handleHamBurger = (id) => {
    const { screen } = this.state;
    if (screen === false) {
      document.getElementById("left-side-menu").style.width = "0";
      document.getElementById("profile").style.width = "0";
      document.getElementById("image").style.width = "0";
      document.getElementById("profilephoto").style.width = "0px";
      document.getElementById("left-side-menu-mini").style.width = "84px";
      document.getElementById("profile-mini").style.width = "80px";
      document.getElementById("image-mini").style.width = "80px";
      document.getElementById("profilephoto-mini").style.width = "70px";
      document.getElementById("side-menu-mini").style.display = "";
      document.getElementById("Qt-logo-mini").style.display = "";
      $('ul#side-menu li.active').css("width", "0");
      $('#mini-sidebar').css("margin-left", "85px");
      this.setState({ screen: true })
      // $('.left-side-menu').hide();
      // $('.left-side-menu-mini').show();
    }
    else if (screen === true) {
      document.getElementById("left-side-menu-mini").style.width = "0";
      document.getElementById("profile-mini").style.width = "0";
      document.getElementById("image-mini").style.width = "0";
      document.getElementById("profilephoto-mini").style.width = "0px";
      document.getElementById("left-side-menu").style.width = "240px";
      document.getElementById("profile").style.width = "225px";
      document.getElementById("image").style.width = "238px";
      document.getElementById("profilephoto").style.width = "70px";
      document.getElementById("side-menu-mini").style.display = "none";
      document.getElementById("Qt-logo-mini").style.display = "none";
      $('ul#side-menu li.active').css("width", "225px");
      $('#mini-sidebar').css("margin-left", "240px");
      this.setState({ screen: false });
      // $('.left-side-menu-mini').hide();
      // $('.left-side-menu').show();
    }
  }
  //lookup
  private async getManager(props): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      this.props.context.msGraphClientFactory.getClient()
        .then((client: MSGraphClient) => {
          // https://graph.microsoft.com/v1.0/me/manager
          let UserEmail = this.props.context.pageContext.legacyPageContext.userEmail;
          return client.api('users?$expand=manager').version('beta').filter(`mail eq '${UserEmail}'`).get((error, response: any, rawResponse?: any) => {
            if (error) {
              // reject(error);
              resolve(null);
              return
            }
            else {
              // resolve(response.value[0].manager.displayName)
              // console.log(response.value[0].manager.displayName)
              this.setState({ ManagerEMail: response.value[0].manager.mail, ManagerTitle: response.value[0].manager.displayName, isManager: response.value[0].manager.displayName ? true : false })
            }
          });
        });
    });
  }

  // private getManager = async () => {
  //   let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
  //   var managerURL = `${this._spservice.rootSite}${this._spservice.rootURL}/_api/web/lists/getbytitle('${this.props.TeamList}')/items?$select=*,Manager/EMail,Manager/Title&$expand=Manager&$filter=StaffId eq ${currentUserId}`
  //   const managerListResult = await this._spservice.get(managerURL);
  //   this.setState({ isManager: managerListResult.data ? managerListResult.data.value : [] })
  // }
  private getProjects = async () => {
    var projectListURL = `${this._spservice.rootSite}${this._spservice.rootURL}/_api/web/lists/getbytitle('${this.props.ProjectList}')/items?$select=*,Title,Client/Title&$filter=IsActive eq 1&$expand=Client&$orderby=Client/Title`;
    const projectListResult = await this._spservice.get(projectListURL);
    ($('#projectName3') as any).select2();
    ($('#projectName2') as any).select2();
    console.log(projectListResult, "result");
    this.setState({ ProjectList: projectListResult.ok ? projectListResult.data.value : [] });
  }
  private getTaskType = async () => {
    var TaskTypeURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.TaskType}')/items`
    const TaskTypeResult = await this._spservice.get(TaskTypeURL);
    this.setState({ TaskType: TaskTypeResult.ok ? TaskTypeResult.data.value : [] });
  }
  //datatable
  private getAllTimesheets = async () => {
    $("#myTabContent").show();
    $("#calendar").hide();
    $(".hide-show-search").show();
    $("#staffTablehide2").show();
    $('#projectName2').val(null).trigger("change");
    $('#projectName3').val(null).trigger("change");
    this.setState({ isPending: false, isCalender: false, project: [], dateRange: 0, errorMessageFromDate: "", errorMessageToDate: "", statusFilter: "", groupColumn: "" });
    await this.loadDefaultDates("0");
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    // var employeeListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items?$select=*,Project/Title,Resource/Title,TaskType/Title&$expand=Project,Resource,TaskType&$filter=ResourceId eq ${currentUserId}`
    var _select = [{ fieldType: null, key: 'Effort,Status,TaskDescription,ProjectID,Date,Resource/Title,TaskType/Title,FieldValuesAsText/TaskDescription' }];
    var _expand = ["Resource,TaskType,FieldValuesAsText"];
    var _filter = `ResourceId eq ${currentUserId}`;
    var _orderby = `Date`;


    // if (this.state.fromDate != "" && this.state.toDate != "") {
    //   _filter += ` and Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment(this.state.toDate).endOf('day').toISOString()}'`;
    // }
    // if (this.state.fromDate && !this.state.toDate) {
    //   _filter += ` and Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment().endOf('day').toISOString()}'`;
    // }
    const employeeListResult = await this._spservice.getPagedListItems(this.props.EmployeeList, _select, _filter, _orderby, _expand);
    destroyDataTable("staffTable");
    destroyDataTable("staffTable1");
    destroyDataTable("myTimesheet-export");
    employeeListResult.reverse();
    let filterdata = [];
    employeeListResult.forEach(element => {
      if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
        filterdata.push(element)
      }
    });
    this.setState({ EmployeeList: filterdata ? filterdata : [] });
    this.callDataTable("staffTable1", this.state.groupColumn);
    this.calcTotalAllEfforts();
    this.calcEfforts();
    this.toggleLoader(false);
  }
  private getPendingTimesheets = async () => {
    // debugger
    $("#myTabContent").show();
    $(".hide-show-search").show();
    $("#calendar").hide();
    $("#staffTablehide1").show(); $('#projectName2').val(null).trigger("change");
    $('#projectName3').val(null).trigger("change");
    this.setState({ isPending: true, isCalender: false, project: [], dateRange: 0, errorMessageFromDate: "", errorMessageToDate: "", statusFilter: "", groupColumn: "" });
    await this.loadDefaultDates("0");
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    // var employeeListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items?$select=*,Project/Title,Resource/Title,TaskType/Title&$expand=Project,Resource,TaskType&$filter=ResourceId eq ${currentUserId}`
    var _select = [{ fieldType: null, key: 'Effort,Status,TaskDescription,ProjectID,Date,Resource/Title,TaskType/Title,FieldValuesAsText/TaskDescription' }];
    var _expand = ["Resource,TaskType,FieldValuesAsText"];
    var _filter = `ResourceId eq ${currentUserId}`;
    var _orderby = `Date`;
    // if (this.state.fromDate != "" && this.state.toDate != "") {
    //   _filter += ` and Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment(this.state.toDate).endOf('day').toISOString()}'`;
    // }
    // if (this.state.fromDate && !this.state.toDate) {
    //   _filter += ` and Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment().endOf('day').toISOString()}'`;
    // }
    const employeeListResult = await this._spservice.getPagedListItems(this.props.EmployeeList, _select, _filter, _orderby, _expand);
    destroyDataTable("staffTable");
    destroyDataTable("staffTable1");
    destroyDataTable("myTimesheet-export");
    var data = employeeListResult?.filter((item) => item.Status.toLowerCase() === "waiting for approval")
    data.reverse();
    let filterdata = [];
    data.forEach(element => {
      if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
        filterdata.push(element)
      }
    });
    this.setState({ EmployeeList: filterdata ? filterdata : [] });
    this.callDataTable("staffTable", this.state.groupColumn);
    this.calcTotalEfforts();
    this.calcEfforts();
    this.toggleLoader(false);
  }
  private handleExcel = () => {
    let filename = `Timesheet_${this._spservice.moment(new Date()).format("YYYY-MM-DD HH:mm:ss")}`;
    download("myTimesheet-export", [], [], filename, "excel", "");
    // const groupBy = this.state.groupColumn;
    // let groupColumn = groupBy ? (parseInt(groupBy)+1) : -1;
    // if(groupColumn >=0 ){ 
    //   if ($('#pending').hasClass('active')){
    //     var colspan1 = 3;
    //     var colspan2 = 1;
    //   }else if ($('#all').hasClass('active')){
    //     var colspan1 = (groupBy=="4")? 4 : 3;
    //     var colspan2 = (groupBy=="4")? 1 : 2;
    //   }          
    //   downloadWithGroupBy("myTimesheet-export", groupColumn, 4, colspan1, colspan2, [], [], filename, "excel", "");
    // }else{
    //   download("myTimesheet-export", [], [], filename, "excel", "");
    // }      
  }
  private callDataTable(tableId: string, groupBy: string) {
    let groupColumn = groupBy ? parseInt(groupBy) : -1;
    if (groupColumn >= 0) {
      if ($('#pending').hasClass('active')) {
        var colspan1 = 2;
        var colspan2 = 2;
      } else if ($('#all').hasClass('active')) {
        var colspan1 = (groupBy == "4") ? 3 : 2;
        var colspan2 = (groupBy == "4") ? 2 : 3;
      }
      applyDataTableGroupBy(tableId, parseInt(groupBy), 3, colspan1, colspan2);
    } else {
      applyDataTable(tableId);
    }
  }
  //filter
  private hideSearch = () => {
    $("#search-filter").css("display", "none");
    $("#search-filter1").css("display", "none");
    $("#hide-search").css("display", "none");
    $("#show-search").css("display", "flex");
  }
  private showSearch = () => {
    //this.setState({toDate: moment().format("YYYY-MM-DD")});
    //$("#toDate").val(moment().format("YYYY-MM-DD"));
    $("#search-filter").css("display", "flex");
    $("#search-filter1").css("display", "flex");
    $("#hide-search").css("display", "flex");
    $("#show-search").css("display", "none");
  }
  private handleSelectedId(id) {
    this.setState({ project: id });
    this.handleSearch();
  }
  private handleDateRange = async (event) => {
    let opt = event.target.value;
    this.setState({ dateRange: opt });
    await this.loadDefaultDates(opt);
    this.handleSearch();
  }
  private handleFromDateChange = (event) => {
    let value = event.target.value
    const now = moment();
    let today = now.format("YYYY-MM-DD");
    let lastMonthFromDate = moment(now).subtract(1, 'months').format('YYYY-MM-DD');
    let lasttwoMonthsFromDate = moment(now).subtract(2, 'months').format('YYYY-MM-DD');
    if (this.state.toDate == today) {
      if (value == lasttwoMonthsFromDate) {
        var defaultdaterange = 0;
      } else if (value == lastMonthFromDate) {
        var defaultdaterange = 1;
      } else if (value == today) {
        var defaultdaterange = 2;
      } else {
        var defaultdaterange = 3;
      }
    } else {
      var defaultdaterange = 3;
    }
    this.setState({
      fromDate: value,
      dateRange: defaultdaterange
    });
    this.handleSearch();
  }
  private handleToDateChange = async (event) => {
    let value = event.target.value
    const now = moment();
    let today = now.format("YYYY-MM-DD");
    let lastMonthFromDate = moment(now).subtract(1, 'months').format('YYYY-MM-DD');
    let lasttwoMonthsFromDate = moment(now).subtract(2, 'months').format('YYYY-MM-DD');
    if (value == today) {
      if (this.state.fromDate == lasttwoMonthsFromDate) {
        var defaultdaterange = 0;
      } else if (this.state.fromDate == lastMonthFromDate) {
        var defaultdaterange = 1;
      } else if (this.state.fromDate == today) {
        var defaultdaterange = 2;
      } else {
        var defaultdaterange = 3;
      }
    } else {
      var defaultdaterange = 3;
    }
    this.setState({
      toDate: value,
      dateRange: defaultdaterange
    });
    this.handleSearch();
  }
  private handleStatus = (event) => {
    let value = event.target.value;
    console.log(value);
    this.setState({ statusFilter: value }, () => {
      this.handleSearch();
    });
  }
  private handleGroupBy = (event) => {
    let value = event.target.value;
    this.setState({ groupColumn: value });
    this.handleSearch();
  }
  private handleSearch = async () => {
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    if (this.state.toDate && !this.state.fromDate) {
      this.setState({
        errorMessageFromDate: "Please enter From Date",
        errorMessageToDate: ""
      });
      return;
    }
    if (moment(this.state.fromDate).isAfter(this.state.toDate)) {
      this.setState({
        errorMessageFromDate: "From Date should be lesser",
        errorMessageToDate: "To Date should be greater"
      });
      return;
    }
    this.setState({
      errorMessageFromDate: "",
      errorMessageToDate: ""
    });
    var filter = "";
    if ((this.state.project).length > 0) {
      this.state.project.forEach((id, index) => {
        if (filter == "") {
          if (index == 0)
            filter = `(ProjectID eq ${id}`;
          else
            filter += ` or ProjectID eq ${id}`
          if (index + 1 == this.state.project.length)
            filter += `)`
        } else {
          if (index == 0)
            filter += ` and (ProjectID eq ${id}`;
          else
            filter += ` or ProjectID eq ${id}`
          if (index + 1 == this.state.project.length)
            filter += `)`
        }
      })
    }
    // if (this.state.fromDate != "" && this.state.toDate != "") {
    //   if (filter == "") {
    //     filter = `Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment(this.state.toDate).endOf('day').toISOString()}'`;
    //   } else {
    //     filter += ` and Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment(this.state.toDate).endOf('day').toISOString()}'`;
    //   }
    // }
    // if (this.state.fromDate && !this.state.toDate) {
    //   if (filter == "") {
    //     filter = `Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment().endOf('day').toISOString()}'`;
    //   } else {
    //     filter += ` and Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment().endOf('day').toISOString()}'`;
    //   }
    // }
    console.log(this.state.statusFilter);
    // if (this.state.statusFilter != "") {
    //   if (filter == "") {
    //     filter = `Status eq '${this.state.statusFilter}'`;
    //   } else {
    //     filter += ` and Status eq '${this.state.statusFilter}'`;
    //   }
    // }
    console.log(filter);
    var _select = [{ fieldType: null, key: 'Effort,Status,TaskDescription,ProjectID,Date,Resource/Title,TaskType/Title,FieldValuesAsText/TaskDescription,Date' }];
    var _expand = ["Resource,TaskType,FieldValuesAsText"];
    var _filter = `${filter === "" ? filter : filter + "and"} ResourceId eq ${currentUserId}`;
    var _orderby = `Date`;
    const searchListResult = await this._spservice.getPagedListItems(this.props.EmployeeList, _select, _filter, _orderby, _expand);
    destroyDataTable("staffTable")
    destroyDataTable("staffTable1")
    destroyDataTable("myTimesheet-export")
    if (this.state.isPending) {
      var data = searchListResult?.filter((item) => item.Status.toLowerCase() === "waiting for approval")
      data.reverse()
      let filterdata = [];
      data.forEach(element => {
        if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
          filterdata.push(element)
        }

      });
      console.log(filterdata);
      this.setState({ EmployeeList: filterdata ? filterdata : [] });
      this.callDataTable("staffTable", this.state.groupColumn);
      this.calcTotalEfforts();
      this.calcEfforts();
    }
    else {
      let filterdata = [];
      if (this.state.statusFilter == "Waiting for Approval") {
        var data = searchListResult?.filter((item) => item.Status.toLowerCase() === "waiting for approval")
        data.reverse()
        data.forEach(element => {
          if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
            filterdata.push(element)
          }

        });
      } else if (this.state.statusFilter == "Approved") {
        var data = searchListResult?.filter((item) => item.Status.toLowerCase() === "approved")
        data.reverse()
        data.forEach(element => {
          if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
            filterdata.push(element)
          }

        });
      } else if (this.state.statusFilter == "Rejected") {
        var data = searchListResult?.filter((item) => item.Status.toLowerCase() === "rejected")
        data.reverse()
        data.forEach(element => {
          if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
            filterdata.push(element)
          }

        });
      } else if (this.state.statusFilter == "Revised") {
        var data = searchListResult?.filter((item) => item.Status.toLowerCase() === "revised")
        data.reverse()
        data.forEach(element => {
          if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
            filterdata.push(element)
          }

        });
      } else {
        searchListResult.reverse()
        searchListResult.forEach(element => {
          if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
            filterdata.push(element)
          }

        });
      }
      this.setState({ EmployeeList: filterdata ? filterdata.reverse() : [] })
      this.callDataTable("staffTable1", this.state.groupColumn);
      this.calcTotalAllEfforts();
      this.calcEfforts();
    }
  }
  private handleClear = async () => {
    $('#projectName2').val(null).trigger("change");
    $('#projectName3').val(null).trigger("change");
    this.setState({ project: [], dateRange: 0, errorMessageFromDate: "", errorMessageToDate: "", statusFilter: "", groupColumn: "" });
    await this.loadDefaultDates("0");
    this.state.isPending ? this.getPendingTimesheets() : this.getAllTimesheets();
  }
  //new timesheet
  private handleNewTimesheet = async () => {
    document.getElementById("newDate").classList.remove("border-danger");
    document.getElementById("selectprojectname").classList.remove("border-danger");
    // document.getElementById("description").classList.remove("border-danger");
    document.getElementById("efforts").classList.remove("border-danger");
    document.getElementById("taskType").classList.remove("border-danger");
    $("#newDate").val(`${this._spservice.moment(new Date()).format("YYYY-MM-DD")}`);
    this.getManager(this.props);
    this.setState({ efforts: "" })
  }
  private handleSubmit = async () => {
    let validateFields: IValidationField[] = [
      { Id: "newDate", Type: "text" },
      { Id: "selectprojectname", Type: "text" },
      // { Id: "description", Type: "text" },
      { Id: "efforts", Type: "text" },
      { Id: "taskType", Type: "select" },
    ];
    const validationResult = this._spservice.utils.validateFields(validateFields, true);
    if (!validationResult.IsValid) {
      return
    }
    var richtext = ($('#description') as any).summernote('code');
    if (($('#description') as any).summernote('isEmpty')) {
      alert("Please enter the Task description");
      return;
    }
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    var date = `${this._spservice.moment($("#newDate").val(), "YYYY-MM-DD").toISOString()}`;
    var project = $("#projectName").val();
    var description = richtext;
    var effort = $("#efforts").val();
    var tasktype = $("#taskType").val();
    var metadata = {
      Date: date,
      ProjectID: project,
      TaskDescription: description,
      Effort: effort,
      TaskTypeId: tasktype,
      Status: "Waiting for Approval",
      Resubmit: false,
      ResourceId: currentUserId
    };
    var postURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items`
    await this._spservice.post(postURL, metadata);
    ($('#custom-modal') as any).modal('hide');
    if (this.state.isCalender) {
      this.calendar();
    } else if (this.state.isPending) {
      this.getPendingTimesheets();
    } else {
      this.getAllTimesheets();
    }
    this.setState({
      successMessage: "Created Successfully",
      efforts: ""
    })
    $("#newDate").val("");
    $("#projectName").val("");
    $("#selectprojectname").val("");
    //$("#description").val("");
    ($('#description') as any).summernote('code', "");
    $("#efforts").val("");
    $("#taskType").val(0);
  }
  //edit timesheet    
  private getItemlist = async (index) => {
    document.getElementById("newDate1").classList.remove("border-danger");
    document.getElementById("selectprojectname1").classList.remove("border-danger");
    // document.getElementById("description1").classList.remove("border-danger");
    document.getElementById("efforts1").classList.remove("border-danger");
    document.getElementById("taskType1").classList.remove("border-danger");

    var employeeListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items(${index})?$select=*,Resource/Title,Resource/EMail&$expand=Resource`;
    const employeeListResult = await this._spservice.get(employeeListURL);

    $("#newDate1").val(`${this._spservice.moment(employeeListResult.data.Date).format("YYYY-MM-DD")}`);
    var listURL = `${this._spservice.rootSite}${this._spservice.rootURL}/_api/web/lists/getbytitle('${this.props.ProjectList}')/items(${employeeListResult.data.ProjectID})?$select=Title,ClientId,Client/Title&$expand=Client`;
    const listResult = await this._spservice.get(listURL);
    if (listResult.ok) {
      let clientname = listResult.data.ClientId ? '[' + listResult.data.Client.Title + '] ' : '';
      console.log(clientname + listResult.data.Title, listResult.data)
      $("#selectprojectname1").val(clientname + listResult.data.Title);
      $("#projectName1").val(employeeListResult.data.ProjectID);
    }
    //$("#description1").val(employeeListResult.data.TaskDescription);
    let status = employeeListResult.data.Status.toLowerCase();
    if (status === "waiting for approval" || status === "revised") {
      ($('#description1') as any).summernote('enable');
    } else {
      ($('#description1') as any).summernote('disable');
    }
    ($('#description1') as any).summernote('code', employeeListResult.data.TaskDescription);
    $("#efforts1").val(employeeListResult.data.Effort);
    $("#comments1").val(employeeListResult.data.Comments);
    $("#taskType1").val(employeeListResult.data.TaskTypeId);
    $("#update").val(employeeListResult.data.Id);
    this.setState({
      status: employeeListResult.data.Status ? employeeListResult.data.Status.toLowerCase() : "", efforts: employeeListResult.data.Effort,
      createdBy: employeeListResult.data.Resource.Title
    })
    this.getManager(this.props);
  }
  private handleUpdate = async () => {
    let validateFields: IValidationField[] = [
      { Id: "newDate1", Type: "text" },
      { Id: "selectprojectname1", Type: "text" },
      { Id: "description1", Type: "text" },
      { Id: "efforts1", Type: "text" },
      { Id: "taskType1", Type: "select" },
    ];
    const validationResult = this._spservice.utils.validateFields(validateFields, true);
    if (!validationResult.IsValid) {
      return
    }
    var richtext = ($('#description1') as any).summernote('code');
    if (($('#description1') as any).summernote('isEmpty')) {
      alert("Please enter the Task description");
      return;
    }
    var date = `${this._spservice.moment($("#newDate1").val(), "YYYY-MM-DD").toISOString()}`;
    var project = $("#projectName1").val();
    var description = richtext;
    var effort = $("#efforts1").val();
    var tasktype = $("#taskType1").val();
    var metadata = {
      Date: date,
      ProjectID: project,
      TaskDescription: description,
      Effort: effort,
      TaskTypeId: tasktype,
      Status: "Waiting for Approval",
      Resubmit: false
    };
    var postURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items(${$("#update").val()})`
    await this._spservice.update(postURL, metadata);
    ($('#custom-modal1') as any).modal('hide');
    this.state.isPending ? this.getPendingTimesheets() : this.getAllTimesheets()
    this.setState({
      successMessage: "Updated Successfully"
    })
  }
  private handleResubmit = async () => {
    let validateFields: IValidationField[] = [
      { Id: "newDate1", Type: "text" },
      { Id: "selectprojectname1", Type: "text" },
      { Id: "description1", Type: "text" },
      { Id: "efforts1", Type: "text" },
      { Id: "taskType1", Type: "select" },
    ];
    const validationResult = this._spservice.utils.validateFields(validateFields, true);
    if (!validationResult.IsValid) {
      return
    }
    var richtext = ($('#description1') as any).summernote('code');
    if (($('#description1') as any).summernote('isEmpty')) {
      alert("Please enter the Task description");
      return;
    }
    var date = `${this._spservice.moment($("#newDate1").val(), "YYYY-MM-DD").toISOString()}`;
    var project = $("#projectName1").val();
    var description = richtext;
    var effort = $("#efforts1").val();
    var tasktype = $("#taskType1").val();
    var metadata = {
      Date: date,
      ProjectID: project,
      TaskDescription: description,
      Effort: effort,
      TaskTypeId: tasktype,
      Status: "Waiting for Approval",
      Resubmit: true
    };
    var postURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items(${$("#update").val()})`
    await this._spservice.update(postURL, metadata);
    ($('#custom-modal1') as any).modal('hide');
    this.state.isPending ? this.getPendingTimesheets() : this.getAllTimesheets()
    this.setState({
      successMessage: "Updated Successfully"
    });
    // , () => {
    //   $('#myAlert').addClass('show');
    //   $('#myAlert').show();
    // })
  }
  private handleDelete = async () => {
    // if (confirm("Are you sure you want to delete?")) {
    var postURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items(${$("#update").val()})`
    await this._spservice.delete(postURL);
    this.handleMail();
    ($('#custom-modal1') as any).modal('hide');
    ($('#custom-modal2') as any).modal('hide');
    this.state.isPending ? this.getPendingTimesheets() : this.getAllTimesheets()
    this.setState({
      successMessage: "Deleted Successfully"
    })
    // , () => {
    //   $('#myAlert').addClass('show');
    //   $('#myAlert').show();
    // })
  }
  handleMail = () => {
    var appweburl = this._spservice.absoluteUrl;
    var urlTemplate = appweburl + "/_api/SP.Utilities.Utility.SendEmail";
    $.ajax({
      contentType: 'application/json',
      url: urlTemplate,
      type: "POST",
      data: JSON.stringify({
        'properties': {
          '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
          // 'From': "gowtham.rajkumar@qantler.com",
          'To': { 'results': [`${this.state.isManager ? this.state.ManagerEMail : null}`] },
          'Body': `<div style="font-size: 15px; font-family: Calibri;"><p>Hi ${this.state.ManagerTitle},<p>
                        <p>It is to inform that the timesheet requested by ${this.state.createdBy} has been deleted due to some reasons.</p>
                        <p style="color: #ff0000 !important;">Note: This is system generated mail, Please do not reply.</p></div>
                   `,
          'Subject': "Reg - Timesheet Request"
        }
      }
      ),
      headers: {
        "Accept": "application/json;odata=verbose",
        "content-type": "application/json;odata=verbose",
        "X-RequestDigest": this._spservice.digest
      },
      success: function (data) {
        console.log("An email was sent.");
      },
      error: function (args) {
        console.log("We had a problem and an email was not sent.");
      }
    });
  }
  private handleInput = (event) => {
    const value = event.target.value
    var pattern = /^\d*(\.\d{0,2})?$/;
    var test = pattern.test(event.target.value)
    if (value === "") {
      this.setState({ efforts: "" })
    }
    else {
      if (test) {
        if (parseInt(value) >= 0 && parseInt(value) <= 24) {
          this.setState({ efforts: value })
        }
      }
    }
  }
  private handleClose = () => {
    ($('#custom-modal') as any).modal('hide');
    $("#selectprojectname").val("");
    $("#projectName").val("");
    //$("#description").val("");
    ($('#description') as any).summernote('code', "");
    $("#efforts").val("");
    $("#taskType").val(0);
    this.setState({ efforts: "" });
  }
  //calculation  
  private calcTotalEfforts() {
    let total = 0;
    // var efforts = document.getElementsByClassName("pendingefforts");
    // for (var i = 0; i < efforts.length; i++){	 
    //   total = +total + (+efforts[i].innerHTML);
    // }
    let efforts = this.state.EmployeeList;
    for (var i = 0; i < efforts.length; i++) {
      total = +total + (+efforts[i].Effort);
    }
    var decimalTimeString: any = total;
    var decimalTime = parseFloat(decimalTimeString);
    decimalTime = decimalTime * 60 * 60;
    var hours: any = Math.floor((decimalTime / (60 * 60)));
    decimalTime = decimalTime - (hours * 60 * 60);
    var minutes: any = Math.floor((decimalTime / 60));
    // decimalTime = decimalTime - (minutes * 60);
    // var seconds: any = Math.round(decimalTime);
    if (hours < 10) {
      hours = "0" + hours;
    }
    if (minutes < 10) {
      minutes = "0" + minutes;
    }
    // if(seconds < 10)
    // {
    //     seconds = "0" + seconds;
    // }
    let totalefforts = hours + ':' + minutes;
    //$('#totalpendingefforts').text(totalefforts); //td total         
    $('#totalpendingefforts').text((total).toFixed(2)); //td total
  }
  private calcTotalAllEfforts() {
    let total = 0;
    // var efforts = document.getElementsByClassName("allefforts");
    // for (var i = 0; i < efforts.length; i++){	 
    //   total = +total + (+efforts[i].innerHTML);
    // }
    let efforts = this.state.EmployeeList;
    for (var i = 0; i < efforts.length; i++) {
      total = +total + (+efforts[i].Effort);
    }
    var decimalTimeString: any = total;
    var decimalTime = parseFloat(decimalTimeString);
    decimalTime = decimalTime * 60 * 60;
    var hours: any = Math.floor((decimalTime / (60 * 60)));
    decimalTime = decimalTime - (hours * 60 * 60);
    var minutes: any = Math.floor((decimalTime / 60));
    // decimalTime = decimalTime - (minutes * 60);
    // var seconds: any = Math.round(decimalTime);
    if (hours < 10) {
      hours = "0" + hours;
    }
    if (minutes < 10) {
      minutes = "0" + minutes;
    }
    // if(seconds < 10)
    // {
    //     seconds = "0" + seconds;
    // }
    let totalefforts = hours + ':' + minutes;
    //$('#totalallefforts').text(totalefforts); //td total      
    $('#totalallefforts').text((total).toFixed(2));
  }
  private calcEfforts() {
    let total = 0;
    // var efforts = document.getElementsByClassName("exportefforts");
    // for (var i = 0; i < efforts.length; i++){	 
    //   total = +total + (+efforts[i].innerHTML);
    // }
    let efforts = this.state.EmployeeList;
    for (var i = 0; i < efforts.length; i++) {
      total = +total + (+efforts[i].Effort);
    }
    var decimalTimeString: any = total;
    var decimalTime = parseFloat(decimalTimeString);
    decimalTime = decimalTime * 60 * 60;
    var hours: any = Math.floor((decimalTime / (60 * 60)));
    decimalTime = decimalTime - (hours * 60 * 60);
    var minutes: any = Math.floor((decimalTime / 60));
    // decimalTime = decimalTime - (minutes * 60);
    // var seconds: any = Math.round(decimalTime);
    if (hours < 10) {
      hours = "0" + hours;
    }
    if (minutes < 10) {
      minutes = "0" + minutes;
    }
    // if(seconds < 10)
    // {
    //     seconds = "0" + seconds;
    // }
    let totalefforts = hours + ':' + minutes;
    //$('#exporttotalefforts').text(totalefforts); //td total       
    $('#exporttotalefforts').text((total).toFixed(2));
  }
  private toggleLoader(IsShow: boolean) {
    if (IsShow)
      $('#loader').css("display", "block");
    else
      $('#loader').css("display", "none");
  }

  private async calendar() {
    $(".hide-show-search").hide();
    $("#staffTablehide1").hide();
    $("#staffTablehide2").hide();
    $("#myTabContent").hide();
    $("#calendar").show();
    $(".fc-button").css("border", "rgb(34, 110, 183)");
    $(".fc-button").css("background-color", "rgb(34, 110, 183)");
    this.setState({ isCalender: true })
    var arr = [];
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    // var employeeListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items?$select=*,Project/Title,Resource/Title,TaskType/Title&$expand=Project,Resource,TaskType&$filter=ResourceId eq ${currentUserId}`
    var _select = [{ fieldType: null, key: '*,Resource/Title,TaskType/Title,FieldValuesAsText/TaskDescription' }];
    var _expand = ["Resource,TaskType,FieldValuesAsText"];
    var _filter = `ResourceId eq ${currentUserId}`;
    var _orderby = `Date`;
    const employeeListResult = await this._spservice.getPagedListItems(this.props.EmployeeList, _select, _filter, _orderby, _expand);
    employeeListResult.map((item) => {
      let Project;
      this.state.ProjectList.map(proitem => {
        if (proitem.Id == item.ProjectID) {
          Project = proitem.Title;
        }
      })
      let Description = item.TaskDescription;
      arr.push({
        title: `${Project}: ${item.Effort}hours`,
        date: moment(item.Date).format("YYYY-MM-DD"),
        extendedProps: {
          effort: item.Effort,
          task: item.TaskType.Title,
          project: Project,
          description: Description
        }
      })
    })
    let User = this.props.context.pageContext.legacyPageContext.userId;
    console.log(User)
    let url = `${this._spservice.rootSite}${this._spservice.rootURL}/LMS/_api/web/lists/getbytitle('${this.props.Leaves}')/items?$filter=AuthorId eq '${User}'`;
    let value = await this._spservice.get(url);
    let data = value.data.value;
    data?.map((item) => {
      console.log(item);
      let leave;
      let color;
      if (item.Status === "Approved") {
        leave = 'Leave'
        color = "#c2c6cc"
      } else {
        leave = 'Waiting for Approval'
        color = "#ff0000"
      }
      // date: `${moment(item.StartDate).format("YYYY-MM-DD").toString()}'-'${moment(item.EndDate).format("YYYY-MM-DD").toString()}`,
      var startDay = new Date(item.StartDate);
      var endDay = new Date(item.EndDate);
      while (startDay <= endDay) {
        arr.push({
          title: leave,
          date: moment(startDay).format("YYYY-MM-DD"),
          overlap: true,
          color: color,
        })
        console.log(startDay);
        var newDate = startDay.setDate(startDay.getDate() + 1);
        startDay = new Date(newDate);
      }
    })
    console.log(arr);
    this.setState({ event: arr }, () => {
      console.log(this.state.event);
    })
  }
  public convertToPlain(html) {

    // Create a new div element
    var tempDivElement = document.createElement("div");

    // Set the HTML content with the given value
    tempDivElement.innerHTML = html;

    // Retrieve the text property of the element 
    return tempDivElement.textContent || tempDivElement.innerText || "";
  }
  handleMouseEnter = (info) => {
    console.log(info.event.extendedProps.description);
    if (info.event.extendedProps.description) {
      $(info.el).tooltip({
        title: '<div style="text-align: left">' + "Project: " + info.event.extendedProps.project + '</br>' + "Task Type: " + info.event.extendedProps.task + '</br>' + "Effort: " + info.event.extendedProps.effort + "hours" + '</br>' + "Description: " + info.event.extendedProps.description + '</div>',
        html: true,
        placement: "left",
        trigger: "hover",
        container: "body"
        });
        $(info.el).tooltip('show');
    }
  };
  public renderEventContent(eventInfo) {
    return (
      <>
        <b>{eventInfo.timeText}</b>
        <i>{eventInfo.event.title}</i>
      </>
    )
  }

  public renderSidebarEvent(event) {
    return (
      <li key={event.id}>
        <b>{formatDate(event.start, { year: 'numeric', month: 'short', day: 'numeric' })}</b>
        <i>{event.title}</i>
      </li>
    )
  }

  handleWeekendsToggle = () => {
    this.setState({
      weekendsVisible: this.state.weekendsVisible
    })
  }

  handleDateSelect = (selectInfo) => {
    let title = prompt('Please enter a new title for your event')
    let calendarApi = selectInfo.view.calendar

    calendarApi.unselect() // clear date selection

    if (title) {
      calendarApi.addEvent({
        id: createEventId(),
        title,
        start: selectInfo.startStr,
        end: selectInfo.endStr,
        allDay: selectInfo.allDay
      })
    }
  }

  handleEventClick = (clickInfo) => {
    if (confirm(`Are you sure you want to delete the event '${clickInfo.event.title}'`)) {
      clickInfo.event.remove()
    }
  }

  handleEvents = (events) => {
    this.setState({
      currentEvents: events
    })

  }
  public render(): React.ReactElement<IMyTimesheetProps> {
    const { EmployeeList, ProjectList, TaskType, status, successMessage, alertMessage, isManager } = this.state;
    return (
      <>
        <div id="loader"></div>
        <div className="content-page">
          <div className="content">
            <div className="container-fluid pl-0 pr-0"></div>
            <div className="row mb-2 add-bar" style={{ paddingBottom: "19px" }}>
              <div className="col-sm-12 col-md-12 float-left ml-4" style={{ paddingRight: "65px" }}>
                <span style={{
                  color: "#000000", right: "10px",
                  top: "58px", fontSize: "20px", cursor: "pointer"
                }} onClick={(e) => this.handleHamBurger(0)}> &#9776; </span>
                <span className='pl-3' style={{ fontSize: "20px", fontWeight: "bold" }}>My Timesheet</span>

                <button onClick={this.handleNewTimesheet} type="button" className="btn btn-primary waves-effect waves-light float-right" data-toggle="modal" data-backdrop="static" data-keyboard="false" data-target="#custom-modal" style={{ backgroundColor: "#226EB7", minHeight: "42px", minWidth: "154px" }}>
                  <i className="fa fa-plus mr-1" /> New Timesheet</button>
                <button onClick={this.handleExcel} type="button" className="btn btn-primary waves-effect waves-light float-right" style={{ backgroundColor: "#226EB7", marginRight: "30px", minHeight: "42px", minWidth: "110px" }}>
                  <i className="fa fa-download mr-2" />Export</button>
              </div>

              <div className="col-sm-12 col-sm-6 col-6 ">
                {/* <div className="app-search-box">
                  <div className="input-group">
                    <input type="search" className="form-control" placeholder="Search..." id="top-search" />
                    <div className="input-group-append">
                      <button className="btn ser-button" type="submit">
                        <i className="fe-search" />
                      </button>
                    </div>
                  </div>
                </div> */}
              </div>
              <div className="col-sm-1"> </div>
            </div>

            <div className='row'>
              <div className='row'>
                <ul className="nav nav-tabs" id="myTab" role="tablist" style={{ margin: "15px 0px 15px 39px", width: "96%" }}>
                  <li className="nav-item" role="presentation">
                    <button className="nav-link active" id="pending-tab" data-toggle="tab" data-target="#pending" type="button" role="tab" aria-controls="pending" aria-selected="true" onClick={this.getPendingTimesheets}>Pending</button>
                  </li>
                  <li className="nav-item" role="presentation">
                    <button className="nav-link" id="all-tab" data-toggle="tab" data-target="#all" type="button" role="tab" aria-controls="all" aria-selected="false" onClick={this.getAllTimesheets}>All</button>
                  </li>
                  <li className="nav-item" role="presentation">
                    <button className="nav-link" id="all-tab" data-toggle="tab" data-target="#all" type="button" role="tab" aria-controls="all" aria-selected="false" onClick={this.calendar.bind(this)} >Calendar View</button>
                  </li>
                  <div className='hide-show-search'>
                    <div className='hide-search' id="hide-search" onClick={this.hideSearch}>Hide Search <i className="arrow-down"></i></div>
                    <div className='hide-search' id="show-search" style={{ display: "none" }} onClick={this.showSearch}>Show Search <i className="arrow-up"></i></div>
                  </div>
                </ul>
              </div>
              <div className="tab-content" id="myTabContent">
                <div className="tab-pane fade show active" id="pending" role="tabpanel" aria-labelledby="home-tab">
                  <div className="row" id="search-filter" style={{ margin: "10px", marginTop: "30px", display: "flex" }}>
                    <div className="col-sm-3 col-md-3">
                      <label htmlFor="position" id="label">Project</label>
                      <select className="form-control place-holder" id="projectName2" multiple={true}>
                        {ProjectList && ProjectList.map((item, index) => (
                          <option value={item.Id}>{`${item.ClientId ? '[' + item.Client.Title + ']' : ''} ${item.Title}`}</option>
                        ))}
                      </select>
                    </div>
                    <div className="col-sm-3 col-md-3">
                      <label htmlFor="position" id="label">Date Range</label>
                      <select className="form-control place-holder" id="daterange"
                        value={this.state.dateRange} onChange={this.handleDateRange}>
                        <option value="1">Today</option>
                        <option value="0">This Month</option>
                        <option value="2">Last Month</option>
                        <option value="4">Last Two Months</option>
                        <option value="5">Last Three Months</option>
                        <option value="3" disabled>Custom</option>
                      </select>
                    </div>
                    <div className="col-sm-3 col-md-3">
                      <label id="label">From Date</label>
                      <input className='form-control datepicker place-holder' type="date" id="fromDate" name="fromDate"
                        value={this.state.fromDate} onChange={this.handleFromDateChange} max={moment().format("YYYY-MM-DD")}></input>
                      <span style={{ color: "#D8000C" }}>{this.state.errorMessageFromDate || null}</span>
                    </div>
                    <div className="col-sm-3 col-md-3">
                      <label id="label">To Date</label>
                      <input className='form-control datepicker place-holder' type="date" id="toDate" name="toDate"
                        value={this.state.toDate} onChange={this.handleToDateChange} max={moment().format("YYYY-MM-DD")}></input>
                      <span style={{ color: "#D8000C" }}>{this.state.errorMessageToDate || null}</span>
                    </div>
                    <div className="col-sm-3 col-md-3 mt-4">
                      <label id="label">Group By</label>
                      <select className="form-control place-holder" id="example-select"
                        value={this.state.groupColumn} onChange={this.handleGroupBy}>
                        <option value="" selected>None</option>
                        <option value="0">Project</option>
                        <option value="2">Task Type</option>
                      </select>
                    </div>
                    <div className="col-sm-6 col-md-6 mt-4"></div>
                    <div className="col-sm-12 col-md-3 mt-2">
                      <div className='row'>
                        <div className="col-xs-12 col-sm-12" style={{ padding: "0px 0px 0px 15px" }}>
                          <button type="button" style={{ width: "50%", height: "96%", color: "#226EB7", border: "1px solid" }} className="btn waves-effect waves-light mt-3 float-end"
                            onClick={this.handleClear}><i className="fa fa-undo mr-1" /> Reset</button>
                        </div>
                        {/* <div className="col-xs-12 col-sm-6" style={{ padding: "0px 15px 0px 8px" }}>
                          <button type="button" style={{ width: "100%", height: "100%", backgroundColor: "#226EB7" }} className="btn btn-primary waves-effect waves-light mt-3" data-dismiss="modal"
                            onClick={this.handleSearch}><i className="fa fa-search mr-1" /> Search</button></div> */}
                      </div>
                    </div>
                  </div>

                  <div className="row" id='staffTablehide1'>
                    <div className="col-xl-12">
                      <div className="card-box bg-trans">
                        <div className="row pt-1">
                          <div className="col-xs-12 col-md-12 pl-4">
                            <span className="pl-4" style={{ fontWeight: "800", color: "#958F8F", fontSize: "13px" }}>{EmployeeList.length ? `${EmployeeList.length} Items Found` : "0 Items Found"}</span>
                            {/* <span className="pl-4" style={{ color: "#958F8F" }}>{EmployeeList.length ? `${EmployeeList.length} Items` : "0 Items"}</span> */}
                          </div>
                        </div>
                        <div className="table-responsive pt-3" style={{ cursor: "pointer", padding: "0px 26px 0px 30px" }}>
                          <table id="staffTable" className="table table-hover product_table">
                            <thead>
                              <tr className="table-active">
                                <th>Project</th>
                                <th>Task Description</th>
                                <th>Task Type</th>
                                <th>Effort (in hours)</th>
                                <th>Date</th>
                              </tr>
                            </thead>
                            <tbody>
                              {EmployeeList && EmployeeList.map((item, index) => (
                                <tr data-toggle="modal" data-target="#custom-modal1"
                                  data-backdrop="static" data-keyboard="false" onClick={() => this.getItemlist(item.Id)} >
                                  <td className="">
                                    {ProjectList && ProjectList.map(proitem => {
                                      if (proitem.Id == item.ProjectID) {
                                        return proitem.ClientId ? ('[' + proitem.Client.Title + '] ' + proitem.Title) : proitem.Title;
                                      }
                                    })
                                    }
                                  </td>
                                  <td className="text-wrap" dangerouslySetInnerHTML={{ __html: item.TaskDescription }}></td>
                                  <td className="">{item.TaskType ? item.TaskType.Title : ""}</td>
                                  <td className="pendingefforts">{item.Effort ? item.Effort : 0}</td>
                                  <td className="text-nowrap">{item.Date ? this._spservice.moment(item.Date).format("YYYY-MM-DD") : ""}</td>
                                </tr>))}
                            </tbody>
                            <tfoot>
                              <tr>
                                <td></td>
                                <td style={{ fontWeight: "bold" }}>{EmployeeList.length ? "Total Efforts: " : ""}</td>
                                <td></td>
                                <td style={{ fontWeight: "bold" }}>{EmployeeList.length ? <span id="totalpendingefforts"></span> : null}</td>
                                <td></td>
                              </tr>
                            </tfoot>
                          </table>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="tab-pane fade show" id="all" role="tabpanel" aria-labelledby="home-tab">
                  <div className="row" id="search-filter1" style={{ margin: "10px", marginTop: "30px", display: "flex" }}>
                    <div className="col-sm-3 col-md-3">
                      <label htmlFor="position" id="label">Project</label>
                      <select className="form-control place-holder multiselect" id="projectName3" multiple={true} onChange={this.handleSearch}>
                        {ProjectList && ProjectList.map((item, index) => (
                          <option value={item.Id}>{`${item.ClientId ? '[' + item.Client.Title + ']' : ''} ${item.Title}`}</option>
                        ))}
                      </select>
                    </div>
                    <div className="col-sm-3 col-md-3">
                      <label htmlFor="position" id="label">Date Range</label>
                      <select className="form-control place-holder" id="daterange1"
                        value={this.state.dateRange} onChange={this.handleDateRange}>
                        <option value="1">Today</option>
                        <option value="0">This Month</option>
                        <option value="2">Last Month</option>
                        <option value="4">Last Two Months</option>
                        <option value="5">Last Three Months</option>
                        <option value="3" disabled>Custom</option>
                      </select>
                    </div>
                    <div className="col-sm-3 col-md-3">
                      <label id="label">From Date</label>
                      <input className='form-control datepicker place-holder' type="date" id="fromDate1" name="fromDate"
                        value={this.state.fromDate} onChange={this.handleFromDateChange} max={moment().format("YYYY-MM-DD")}></input>
                      <span style={{ color: "#D8000C" }}>{this.state.errorMessageFromDate || null}</span>
                    </div>
                    <div className="col-sm-3 col-md-3">
                      <label id="label">To Date</label>
                      <input className='form-control datepicker place-holder' type="date" id="toDate" name="toDate"
                        value={this.state.toDate} onChange={this.handleToDateChange} max={moment().format("YYYY-MM-DD")}></input>
                      <span style={{ color: "#D8000C" }}>{this.state.errorMessageToDate || null}</span>
                    </div>
                    <div className="col-sm-3 col-md-3 mt-4">
                      <label htmlFor="position" id="label">Status</label>
                      <select className="form-control place-holder" id="example-select"
                        value={this.state.statusFilter} onChange={this.handleStatus}>
                        <option value=""></option>
                        <option value="Waiting for Approval">Waiting for Approval</option>
                        <option value="Approved">Approved</option>
                        <option value="Rejected">Rejected</option>
                        <option value="Revised">Revised</option>
                      </select>
                    </div>
                    <div className="col-sm-3 col-md-3 mt-4">
                      <label id="label">Group By</label>
                      <select className="form-control place-holder" id="example-select"
                        value={this.state.groupColumn} onChange={this.handleGroupBy}>
                        <option value="" selected>None</option>
                        <option value="0">Project</option>
                        <option value="2">Task Type</option>
                        <option value="4">Status</option>
                      </select>
                    </div>
                    <div className="col-sm-3 col-md-3 mt-4"></div>
                    <div className="col-xs-12 col-sm-3 mt-2">
                      <div className='row'>
                        <div className="col-xs-12 col-sm-12" style={{ padding: "0px 0px 0px 15px" }}>
                          <button type="button" style={{ width: "50%", height: "96%", color: "#226EB7", border: "1px solid" }} className="btn waves-effect waves-light mt-3 float-end"
                            onClick={this.handleClear}><i className="fa fa-undo mr-1" /> Reset</button>
                        </div>
                        {/* <div className="col-xs-12 col-sm-6" style={{ padding: "0px 15px 0px 8px" }}>
                          <button type="button" style={{ width: "100%", height: "100%", backgroundColor: "#226EB7" }} className="btn btn-primary waves-effect waves-light mt-3" data-dismiss="modal"
                            onClick={this.handleSearch}><i className="fa fa-search mr-1" /> Search</button></div> */}
                      </div>
                    </div>
                  </div>

                  <div className="row" id='staffTablehide2'>
                    <div className="col-xl-12">
                      <div className="card-box bg-trans">
                        <div className="row pt-1">
                          <div className="col-xs-12 col-md-12 pl-4">
                            <span className="pl-4" style={{ fontWeight: "800", color: "#958F8F", fontSize: "13px" }}>{EmployeeList.length ? `${EmployeeList.length} Items Found` : "0 Items Found"}</span>
                            {/* <span className="pl-4" style={{ color: "#958F8F" }}>{EmployeeList.length ? `${EmployeeList.length} Items` : "0 Items"}</span> */}
                          </div>
                        </div>
                        <div className="table-responsive pt-3" style={{ cursor: "pointer", padding: "0px 26px 0px 30px" }}>
                          <table id="staffTable1" className="table table-hover product_table">
                            <thead>
                              <tr className="table-active">
                                <th>Project</th>
                                <th>Task Description</th>
                                <th>Task Type</th>
                                <th>Effort (in hours)</th>
                                <th>Status</th>
                                <th>Date</th>
                              </tr>
                            </thead>
                            <tbody>
                              {EmployeeList && EmployeeList.map((item, index) => (
                                <tr data-toggle="modal" data-backdrop="static" data-keyboard="false"
                                  onClick={() => this.getItemlist(item.Id)} data-target="#custom-modal1">
                                  <td className="">
                                    {ProjectList && ProjectList.map(proitem => {
                                      if (proitem.Id == item.ProjectID) {
                                        return proitem.ClientId ? ('[' + proitem.Client.Title + '] ' + proitem.Title) : proitem.Title;
                                      }
                                    })
                                    }
                                  </td>
                                  <td className="text-wrap" dangerouslySetInnerHTML={{ __html: item.TaskDescription }}></td>
                                  <td className="">{item.TaskType ? item.TaskType.Title : ""}</td>
                                  <td className="allefforts">{item.Effort ? item.Effort : 0}</td>
                                  <td className="">{item.Status || ""}</td>
                                  <td className="text-nowrap">{item.Date ? this._spservice.moment(item.Date).format("YYYY-MM-DD") : ""}</td>
                                </tr>
                              ))
                              }
                            </tbody>
                            <tfoot>
                              <tr>
                                <td></td>
                                <td style={{ fontWeight: "bold" }}>{EmployeeList.length ? "Total Efforts: " : ""}</td>
                                <td></td>
                                <td style={{ fontWeight: "bold" }}>{EmployeeList.length ? <span id="totalallefforts"></span> : null}</td>
                                <td></td>
                                <td></td>
                              </tr>
                            </tfoot>
                          </table>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

              </div>
            </div>

          </div>
        </div>

        <div className="modal left fade" id="custom-modal" tabIndex={-1} role="dialog" aria-hidden="true">
          <div className="modal-dialog modal-dialog-centered">
            <div className="modal-content">
              <div className="modal-header bg-light">
                <div className="row">
                  <div className="col-2 col-md-2 col-sm-2" style={{ padding: "3px" }}>
                    <span className="close" style={{ textAlign: "right", fontSize: "24px !important" }} onClick={this.handleClose} aria-hidden="true">×</span>
                  </div>
                  <div className="col-10 col-md-10 col-sm-10">
                    <h4 className="modal-title" id="myCenterModalLabel">New Timesheet</h4>
                  </div>
                </div>
              </div>
              <div className="modal-body p-4">
                <form id="myform">
                  <div className="form-group">
                    <label htmlFor="name">Date</label>
                    <input max={`${this._spservice.moment(new Date()).format("YYYY-MM-DD")}`} className='form-control datepicker' type="date" id="newDate" name="newDate"
                    ></input>
                  </div>
                  <div className="form-group">
                    <label htmlFor="name">Project</label>
                    <input className="form-control" type="text" id="selectprojectname" list="projectsandclients" autoComplete='false' />
                    <input className="form-control" type="hidden" id="projectName" name="projectName" />
                    <datalist id="projectsandclients">
                      {ProjectList && ProjectList.map((item, index) => (
                        <option data-id={item.Id} id={item.Title} value={`${item.ClientId ? '[' + item.Client.Title + ']' : ''} ${item.Title}`}></option>
                      ))}
                    </datalist>
                  </div>
                  <div className="form-group" id='summernotedescription'>
                    <label htmlFor="position">Task Description</label>
                    <div className="summernote" id="description"></div>
                    {/* <textarea className="form-control" id="description" rows={4}/>  */}
                  </div>
                  <div className="form-group">
                    <label htmlFor="category">Effort (in hours)</label>
                    <input type="text" className="form-control" id="efforts" name="efforts" value={this.state.efforts} onChange={this.handleInput}></input>
                  </div>
                  <div className="form-group">
                    <label htmlFor="category">Task Type</label>
                    <select className="form-control" id="taskType">
                      <option value={0}>-- Select --</option>
                      {TaskType && TaskType.map((item, index) => (
                        <option value={item.Id}>{item.Title}</option>
                      ))}
                    </select>
                  </div>
                  {isManager ? null :
                    <div className="form-group">
                      <div className="alert alert-danger">
                        <strong>Warning!</strong> No manager is assigned to you. Please contact HR.
                      </div>
                    </div>
                  }
                </form>
              </div>
              <div className="modal-footer text-center display-block p-4">
                <button type="button" className="btn btn-secondary right-button"
                  onClick={this.handleClose}>Close</button>
                <button type="button" style={{ backgroundColor: "#226EB7" }} className="btn btn-primary left-button"
                  onClick={this.handleSubmit}>Submit</button>

              </div>
            </div>
          </div>
        </div >

        <div className="modal left fade" id="custom-modal1" tabIndex={-1} role="dialog" aria-hidden="true">
          <div className="modal-dialog modal-dialog-centered">
            <div className="modal-content">
              <div className="modal-header bg-light">
                <div className="row">
                  <div className="col-2 col-md-2 col-sm-2" style={{ padding: "3px" }}>
                    <span style={{ textAlign: "right", fontSize: "24px !important" }} className="close" data-dismiss="modal" aria-hidden="true" >×</span>
                  </div>
                  <div className="col-10 col-md-10 col-sm-10">
                    <h4 className="modal-title" id="myCenterModalLabel">Edit Timesheet</h4>
                  </div>
                </div>
              </div>
              <div className="modal-body p-4">
                <form id="myform1">
                  <div className="form-group">
                    <label htmlFor="name">Date</label>
                    <input max={`${this._spservice.moment(new Date()).format("YYYY-MM-DD")}`} className='form-control datepicker' disabled={status === "waiting for approval" || status === "revised" ? false : true} type="date" id="newDate1" name="newDate1"
                    ></input>
                  </div>
                  <div className="form-group">
                    <label htmlFor="name">Project</label>
                    <input className="form-control" type="text" id="selectprojectname1" list="projectsandclients1" autoComplete='false'
                      disabled={status === "waiting for approval" || status === "revised" ? false : true} />
                    <input className="form-control" type="hidden" id="projectName1" name="projectName1" />
                    <datalist id="projectsandclients1">
                      {ProjectList && ProjectList.map((item, index) => (
                        <option data-id={item.Id} id={item.Title} value={`${item.ClientId ? '[' + item.Client.Title + ']' : ''} ${item.Title}`}></option>
                      ))}
                    </datalist>
                  </div>
                  <div className="form-group" id='summernotedescription'>
                    <label htmlFor="position">Task Description</label>
                    <div className="summernote" id="description1"></div>
                    {/* <textarea className="form-control" id="description1" name="description1" disabled={status === "waiting for approval" || status === "revised" ? false : true} rows={4} /> */}
                  </div>
                  <div className="form-group">
                    <label htmlFor="category">Effort (in hours)</label>
                    <input type="text" className="form-control" id="efforts1" name="efforts1" value={this.state.efforts} onChange={this.handleInput} disabled={status === "waiting for approval" || status === "revised" ? false : true}></input>
                  </div>
                  <div className="form-group">
                    <label htmlFor="category">Task Type</label>
                    <select className="form-control" id="taskType1" disabled={status === "waiting for approval" || status === "revised" ? false : true}>
                      <option value={0}>-- Select --</option>
                      {TaskType && TaskType.map((item, index) => (
                        <option value={item.Id}>{item.Title}</option>
                      ))}
                    </select>
                  </div>
                  {status === "waiting for approval" ? null :
                    <div className="form-group">
                      <label htmlFor="position">Comments</label>
                      <textarea className="form-control" id="comments1" name="comments1" disabled={true} rows={4} />
                    </div>
                  }
                </form>
              </div>
              <div className="modal-footer text-center display-block p-4">
                <button type="button" id='delete' className={`${status === "waiting for approval" || status === "revised" ? "" : "d-none"} btn btn-danger`} style={{ width: "31%" }} data-toggle="modal" data-target="#custom-modal2">Delete</button>
                {status === "waiting for approval" || status === "revised" ?
                  <button type="button" className="btn btn-secondary" style={{ width: "31%" }} data-dismiss="modal">Close</button>
                  :
                  <button type="button" className="btn btn-secondary" style={{ width: "100%" }} data-dismiss="modal">Close</button>}
                <button type="button" id="update" className={`${status === "waiting for approval" ? "" : "d-none"} btn btn-primary`} style={{ width: "32%", backgroundColor: "#226EB7" }} onClick={this.handleUpdate}>Update</button>
                <button type="button" id="update" className={`${status === "revised" ? "" : "d-none"} btn btn-primary`} style={{ width: "32%", backgroundColor: "#226EB7" }} onClick={this.handleResubmit}>Resubmit</button>
              </div>
            </div>
          </div>
        </div>

        <div className="modal fade" id="custom-modal2" tabIndex={-1} role="dialog">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" style={{ margin: "0px" }}>Delete</h5>
                <span className="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true" style={{ marginLeft: "80px" }}>&times;</span>
                </span>
              </div>
              <div className="modal-body">
                <p>Are you sure want to delete?</p>
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-primary" style={{ backgroundColor: "#226EB7" }} onClick={this.handleDelete}>OK</button>
                <button type="button" className="btn btn-secondary" data-dismiss="modal">Cancel</button>
              </div>
            </div>
          </div>
        </div>

        {/* <div id="myAlert" className="alert alert-success fade" style={{ position: "fixed", top: "60px", right: "25px", zIndex: "1111" }}>
          <strong>Note!</strong> {successMessage || "unknown error occured"}
          <button id="myBtn" type="button" className="ml-2 mb-1 close" data-dismiss="alert" aria-label="Close">
            <span aria-hidden="true" style={{ paddingLeft: "130px" }}>&times;</span>
          </button>
        </div> */}

        <div className="modal fade" id="custom-modal3" tabIndex={-1} role="dialog">
          <div className="modal-dialog" role="document">
            <div className="modal-content">
              <div className="modal-body">
                <p>{alertMessage}</p>
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-primary btn-sm" style={{ backgroundColor: "#226EB7" }} data-dismiss="modal">OK</button>
              </div>
            </div>
          </div>
        </div>

        <table id="myTimesheet-export" className="display nowrap d-none" style={{ width: "100%" }}>
          <thead>
            <tr>
              <th>Date</th>
              <th>Project</th>
              <th>Task Description</th>
              <th>Task Type</th>
              <th>Effort (in hours)</th>
              <th>{this.state.isPending ? "" : "Status"}</th>
            </tr>
          </thead>
          <tbody>
            {EmployeeList && EmployeeList.map((item, i) => {
              return (
                <tr key={i}>
                  <td>{item.Date ? this._spservice.moment(item.Date).format("YYYY-MM-DD") : ""}</td>
                  <td>
                    {ProjectList && ProjectList.map(proitem => {
                      if (proitem.Id == item.ProjectID) {
                        return proitem.ClientId ? ('[' + proitem.Client.Title + '] ' + proitem.Title) : proitem.Title;
                      }
                    })
                    }
                  </td>
                  <td dangerouslySetInnerHTML={{ __html: item.TaskDescription }}></td>
                  <td className="">{item.TaskType ? item.TaskType.Title : ""}</td>
                  <td className="exportefforts">{item.Effort ? item.Effort : 0}</td>
                  <td>{this.state.isPending ? "" : item.Status}</td>
                </tr>);
            })}
          </tbody>
          <tfoot>
            <tr>
              <td></td>
              <td></td>
              <td>Total Efforts: </td>
              <td></td>
              <td><span id="exporttotalefforts"></span></td>
              <td></td>
            </tr>
          </tfoot>
        </table>

        <div id='calendar' style={{ padding: "2.5%" }}>
          <div className='demo-app'>
            <div className='demo-app-main'>
              <FullCalendar
                plugins={[dayGridPlugin]}
                headerToolbar={{
                  left: 'today',
                  center: 'title',
                  right: 'prev,next'
                }}
                buttonText={{
                  today: 'Today'
                }}
                initialView='dayGridMonth'
                eventMouseEnter={this.handleMouseEnter}
                editable={true}
                selectable={true}
                selectMirror={true}
                dayMaxEvents={true}
                events={this.state.event}
              />
            </div>
          </div>
        </div>
      </>
    );
  }
}
