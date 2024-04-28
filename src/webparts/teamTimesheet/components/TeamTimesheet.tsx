import * as React from 'react';
import './TeamTimesheet.module.scss';
import * as moment from 'moment';
import * as $ from 'jquery';
import { ITeamTimesheetProps } from './ITeamTimesheetProps';
import SPService from '../../../shared/services/SPService';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { download, downloadWithGroupBy } from '../../../shared/utils/datatable-export';
import { IValidationField } from '../../../shared/models/IValidationField';
import { IBasePickerStyles } from "office-ui-fabric-react/lib/Pickers";
import { applyDataTable, destroyDataTable, applyDataTableGroupBy } from '../../../shared/utils/datatable';
import "select2/dist/css/select2.min.css";
import "select2/dist/js/select2.min.js";
import 'jquery-ui-dist/jquery-ui.min.js';
import 'jquery-ui-dist/jquery-ui.min.css';
require("../../../shared/js/summernote-bs4.js");
import { initializeIcons } from '@fluentui/react/lib/Icons';
initializeIcons();
import FullCalendar, { formatDate } from '@fullcalendar/react' // must go before plugins
import dayGridPlugin from '@fullcalendar/daygrid' // a plugin!
import { MSGraphClient } from '@microsoft/sp-http';

let tooltipInstance = null;

export interface ITeamTimesheetState {
  TeamListData: any[];
  groupName: string;
  ProjectList: any[];
  project: any;
  dateRange: any;
  fromDate: any;
  toDate: any;
  errorMessageFromDate: string;
  errorMessageToDate: string;
  PeopleId: any;
  TaskType: any[];
  status: string;
  statusFilter: string;
  successMessage: string;
  alertMessage: string;
  screen: boolean;
  isPending: boolean;
  groupColumn: any;
  isAdmin: boolean;
  isHR: boolean;
  event: any,
  staff: any[];
  allUser: any[];
  userMail: any;
}

export default class TeamTimesheet extends React.Component<ITeamTimesheetProps, ITeamTimesheetState> {
  private _spservice: SPService;
  protected ppl;
  constructor(props) {
    super(props);
    this._spservice = new SPService(this.props.context);
    this.state = {
      TeamListData: [],
      groupName: "",
      ProjectList: [],
      project: [],
      dateRange: 0,
      fromDate: "",
      toDate: "",
      errorMessageFromDate: "",
      errorMessageToDate: "",
      PeopleId: 0,
      TaskType: [],
      status: "",
      statusFilter: "",
      successMessage: "",
      alertMessage: "",
      screen: false,
      isPending: true,
      groupColumn: "",
      isAdmin: false,
      isHR: false,
      event: [],
      staff: [],
      allUser: [],
      userMail: this.props.context.pageContext.legacyPageContext.userEmail,
    }
  }

  private pickerStylesSingle: Partial<IBasePickerStyles> = {
    root: {},
    input: {
      borderColor: "#ced4da !important",
      height: "48px",
      border: "1px solid #ced4da",
      borderRadius: "4px",
      "&:hover": {
        borderColor: "none !important"
      },
      "&:focus": {
        borderRadius: "4px !important",
        backgroundColor: "#fff",
        borderColor: "#ffffff !important",
        outline: 0,
        boxShadow: "0 0 0 0.3rem rgb(0 123 255 / 25%)"
      },
    },
    text: {
      height: "48px",
      border: "none",
      borderRadius: "4px",
      background: "#ffffff"

      // height: "48px",
      // border: "1px solid #ced4da",
      // borderRadius: "4px"
    },
  }

  componentDidMount = async () => {
    this.toggleLoader(true);
    $("#calendar").hide();
    let _this = this;
    await this.checkPermission();
    await this.getProjects();
    await this.getTaskType();
    await this.getAllUser();
    if ($('#pending').hasClass('active')) {
      this.getPendingTimesheets();
    } else if ($('#all').hasClass('active')) {
      await this.getManager();
    }
    $(document).on('change', '#projectName2', function () {
      let id = $('#projectName2').val();
      _this.handleSelectedId(id);
      _this.handleSearch();
    });
    $(document).on('change', '#projectName3', function () {
      let id = $('#projectName3').val();
      _this.handleSelectedId(id);
      _this.handleSearch();
    });
    ($('#description1') as any).summernote({
      addDefaultFonts: false,
      height: 200,
      toolbar: false,
    });
  }
  private async loadDefaultDates(option) {
    let opt = option;
    //let opt = $('#daterange').val();
    //let opt = this.state.dateRange;
    const now = moment();
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
  private async checkPermission() {
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    var adminURL = `${this._spservice.absoluteUrl}/_api/web/GetUserById('${currentUserId}')/Groups?$filter=Title eq '${this.props.adminGroup}'`;
    const adminResult = await this._spservice.get(adminURL);
    var hrURL = `${this._spservice.absoluteUrl}/_api/web/GetUserById('${currentUserId}')/Groups?$filter=Title eq '${this.props.hrGroup}'`;
    const hrResult = await this._spservice.get(hrURL);
    this.setState({
      isAdmin: adminResult.data.value.length > 0 ? true : false,
      isHR: hrResult.data.value.length > 0 ? true : false
    });
  }

  //lookup  
  private getTaskType = async () => {
    var TaskTypeURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.TaskType}')/items`
    const TaskTypeResult = await this._spservice.get(TaskTypeURL);
    this.setState({ TaskType: TaskTypeResult.ok ? TaskTypeResult.data.value : [] });
  }
  private getProjects = async () => {
    var projectListURL = `${this._spservice.rootSite}${this._spservice.rootURL}/_api/web/lists/getbytitle('${this.props.ProjectList}')/items?$select=*,Title,Client/Title&$filter=IsActive eq 1&$expand=Client&$orderby=Client/Title`
    const projectListResult = await this._spservice.get(projectListURL);
    ($('#projectName2') as any).select2();
    ($('#projectName3') as any).select2();
    this.setState({ ProjectList: projectListResult.ok ? projectListResult.data.value : [] });
  }
  private async getAllUser() {
    var allUser = `${this._spservice.rootSite}${this._spservice.rootURL}/_api/web/siteusers?$top=4999`
    const allUserValue = await this._spservice.get(allUser);
    let getUser = allUserValue.data.value;
    this.setState({ allUser: getUser });
  }
  //datatable
  private async getManager(): Promise<any> {
    new Promise<any>((resolve, reject) => {
      this.props.context.msGraphClientFactory.getClient()
        .then(async (client: MSGraphClient) => {
          // https://graph.microsoft.com/v1.0/me/manager
          client.api('users').version('beta').filter(`mail eq '${this.state.userMail}'`).get((error, response: any, rawResponse?: any) => {
            let managetId = response.value[0].id
            this.getStaff(managetId);
            if (error) {
              // reject(error);.filter(`manager.mail eq '${UserEmail}'`)
              resolve(null);
              return
            }
            else {
              // resolve(response.value[0].manager.displayName)
            }
          });
        });
      this.toggleLoader(false);
    });
  }
  private async getStaff(managetId): Promise<any> {
    new Promise<any>((resolve, reject) => {
      this.props.context.msGraphClientFactory.getClient()
        .then(async (client: MSGraphClient) => {
          // https://graph.microsoft.com/v1.0/me/manager
          client.api(`users/${managetId}/directReports`).version('beta').get((error, response: any, rawResponse?: any) => {
            if (error) {
              // reject(error);.filter(`manager.mail eq '${UserEmail}'`)
              resolve(null);
              return
            }
            else {
              // resolve(response.value[0].manager.displayName)
              let staffArr = [];
              // let currentUser = "nandhini.thiraviyam@qantler.com"
              response.value?.map((element) => {
                this.state.allUser?.map((value) => {
                  if (value.Email === element.mail) {
                    staffArr.push({ Id: value.Id })
                  }
                })
              })
              this.setState({ staff: staffArr })
              this.getAllTimesheets();
            }
          });
        });
      this.toggleLoader(false);
    });
    return true;
  }
  private getPendingTimesheets = async () => {
    $(".hide-show-search").show();
    $('#myTabContent').show();
    $("#calendar").hide();
    $("#Pendingshow").show();
    $('#projectName2').val(null).trigger("change");
    $('#projectName3').val(null).trigger("change");
    this.setState({ isPending: true, project: [], dateRange: 0, errorMessageFromDate: "", errorMessageToDate: "", PeopleId: 0, statusFilter: "", groupColumn: "" });
    await this.loadDefaultDates("0");
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    var filter = "";
    if (!this.state.isAdmin && !this.state.isHR) {

      this.state.staff.forEach((ele, index) => {
        if (filter == "") {
          if (index == 0)
            filter = `ResourceId eq ${ele.Id}`;
          else
            filter += ` or ResourceId eq ${ele.Id}`
        } else {
          if (index == 0)
            filter += ` and ResourceId eq ${ele.Id}`;
          else
            filter += ` or ResourceId eq ${ele.Id}`
        }
      });
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
    // var teamListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items?$select=*,Project/Title,Resource/Title,TaskType/Title&$expand=Project,Resource,TaskType&$filter=${filter}`
    var _select = [{ fieldType: null, key: '*,Resource/Title,Resource/EMail,TaskType/Title,FieldValuesAsText/TaskDescription' }];
    var _expand = ["Resource,TaskType,FieldValuesAsText"];
    var _orderby = `Date`;
    var _filter = `${filter}`;
    const teamListResult1 = await this._spservice.getPagedListItems(this.props.EmployeeList, _select, _filter, _orderby, _expand);
    destroyDataTable("staffTable")
    destroyDataTable("staffTable1")
    destroyDataTable("teamTimesheet-export")
    var data = teamListResult1?.filter((item) => item.Status.toLowerCase() === "waiting for approval")
    data.reverse();
    let filterdata = [];
    data.forEach(element => {
      if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
        filterdata.push(element)
      }
    });
    this.setState({ TeamListData: filterdata ? filterdata : [] });
    this.callDataTable("staffTable", this.state.groupColumn);
    this.calcTotalEfforts();
    this.calcEfforts();
  }
  private getAllTimesheets = async () => {
    $(".hide-show-search").show();
    $("#calendar").hide();
    $("#Allshow").show();
    $('#myTabContent').show();
    $('#projectName2').val(null).trigger("change");
    $('#projectName3').val(null).trigger("change");
    this.setState({ isPending: false, project: [], dateRange: 0, errorMessageFromDate: "", errorMessageToDate: "", PeopleId: 0, statusFilter: "", groupColumn: "" });
    await this.loadDefaultDates("0");
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    if (this.state.isAdmin || this.state.isHR) {
      this.getTeamListData();
    } else if (!this.state.isAdmin && !this.state.isHR) {
      this.getTeamListData(this.state.staff);
    }
    this.toggleLoader(false);
  }
  private getTeamListData = async (staff?) => {
    var filter = "";
    if (!this.state.isAdmin && !this.state.isHR) {
      staff?.forEach((ele, index) => {
        if (filter == "") {
          if (index == 0)
            filter = `ResourceId eq ${ele.Id}`;
          else
            filter += ` or ResourceId eq ${ele.Id}`
        } else {
          if (index == 0)
            filter += ` and ResourceId eq ${ele.Id}`;
          else
            filter += ` or ResourceId eq ${ele.Id}`
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
    // var teamListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items?$select=*,Project/Title,Resource/Title,TaskType/Title&$expand=Project,Resource,TaskType&$filter=${filter}`
    if (filter !== "") {
      var _select = [{ fieldType: null, key: '*,Resource/Title,TaskType/Title,FieldValuesAsText/TaskDescription' }];
      var _expand = ["Resource,TaskType,FieldValuesAsText"];
      var _orderby = `Date`;
      var _filter = `${filter}`;
      const teamListResult = await this._spservice.getPagedListItems(this.props.EmployeeList, _select, _filter, _orderby, _expand);
      destroyDataTable("staffTable");
      destroyDataTable("staffTable1");
      destroyDataTable("teamTimesheet-export");
      teamListResult.reverse();
      let filterdata = [];
      teamListResult.forEach(element => {
        if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
          filterdata.push(element)
        }
      });
      this.setState({ TeamListData: filterdata ? filterdata : [] });
    }
    ($('#projectName3') as any).select2();
    this.callDataTable("staffTable1", this.state.groupColumn);
    this.calcTotalAllEfforts();
    this.calcEfforts();
  }
  private handleExcel = () => {
    let filename = `MyTeamTimesheet_${this._spservice.moment(new Date()).format("YYYY-MM-DD HH:mm:ss")}`
    download("teamTimesheet-export", [], [], filename, "excel", "");
    // const groupBy = this.state.groupColumn;
    // let groupColumn = groupBy ? (parseInt(groupBy)+1) : -1;
    // if(groupColumn >=0 ){ 
    //   if ($('#pending').hasClass('active')){
    //     var colspan1 = 3;
    //     var colspan2 = 1;
    //   }else if ($('#all').hasClass('active')){
    //     var colspan1 = (groupBy=="5")? 5 : 4;
    //     var colspan2 = (groupBy=="5")? 1 : 2;
    //   }          
    //   downloadWithGroupBy("teamTimesheet-export", groupColumn, 5, colspan1, colspan2, [], [], filename, "excel", "");
    // }else{
    //   download("teamTimesheet-export", [], [], filename, "excel", "");
    // }      
  }
  private callDataTable(tableId: string, groupBy: string) {
    if ($('#pending').hasClass('active')) {
      var colspan1 = 3;
      var colspan2 = 2;
    } else if ($('#all').hasClass('active')) {
      var colspan1 = (groupBy == "5") ? 4 : 3;
      var colspan2 = (groupBy == "5") ? 2 : 3;
    }
    let groupColumn = groupBy ? parseInt(groupBy) : -1;
    if (groupColumn >= 0) {
      applyDataTableGroupBy(tableId, parseInt(groupBy), 4, colspan1, colspan2);
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
  private getPeople = async (items) => {
    let value = items[0]?.id
    if (items.length > 0) {
      this.setState({ PeopleId: value }, () => {
        this.handleSearch();
      })
    }
    else {
      this.setState({ PeopleId: 0 })
    }
  }
  private handleSelectedId(id) {
    this.setState({ project: id });
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
  private handleToDateChange = (event) => {
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
    // if (!this.state.toDate && !this.state.fromDate && !this.state.project && this.state.PeopleId === 0) {
    //   this.setState({
    //     alertMessage: "Please select atleast one field"
    //   }, () => {
    //     $('#custom-modal3').modal('show');
    //   })
    //   return;
    // }

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
    if (this.state.PeopleId !== 0) {
      var isStaff = this.state.staff.filter((item) => (item.Id === this.state.PeopleId));
      if (!this.state.isAdmin && !this.state.isHR) {
        if (filter == "") {
          filter = isStaff.length > 0 ? `ResourceId eq ${this.state.PeopleId}` : "ResourceId eq null";
        } else {
          filter += isStaff.length > 0 ? ` and ResourceId eq ${this.state.PeopleId}` : " and ResourceId eq null";
        }
      } else {
        if (filter == "") {
          filter = `ResourceId eq ${this.state.PeopleId}`;
        } else {
          filter += ` and ResourceId eq ${this.state.PeopleId}`;
        }
      }
    } else {
      //default filter for manager
      if (!this.state.isAdmin && !this.state.isHR) {
        this.state.staff.forEach((ele, index) => {
          if (index == 0)
            filter = `(ResourceId eq ${ele.Id}`;
          else
            filter += ` or ResourceId eq ${ele.Id}`
          if (index + 1 == this.state.staff.length) {
            filter += `)`
          }
        })
      }
    }
    if ((this.state.project).length > 0) {
      debugger
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
    // if (this.state.fromDate !== "" && this.state.toDate !== "") {
    //   if (filter == "") {
    //     filter = `Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment(this.state.toDate).endOf('day').toISOString()}'`
    //   } else {
    //     filter += ` and Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment(this.state.toDate).endOf('day').toISOString()}'`
    //   }
    // }
    // if (this.state.fromDate !== "" && this.state.toDate === "") {
    //   if (filter == "") {
    //     filter = `Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment().endOf('day').toISOString()}'`
    //   } else {
    //     filter += ` and Date ge datetime'${moment(this.state.fromDate).startOf('day').toISOString()}' and Date le datetime'${moment().endOf('day').toISOString()}'`
    //   }
    // }
    if (this.state.statusFilter !== "") {
      if (filter == "") {
        filter = `Status eq '${this.state.statusFilter}'`;
      } else {
        filter += ` and Status eq '${this.state.statusFilter}'`;
      }
    }
    // var searchListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items?$select=*,Project/Title,Resource/Title,TaskType/Title&$expand=Project,Resource,TaskType&$filter=${filter}`
    var _select = [{ fieldType: null, key: '*,Resource/Title,TaskType/Title,FieldValuesAsText/TaskDescription' }];
    var _expand = ["Resource,TaskType,FieldValuesAsText"];
    var _filter = `${filter}`;
    var _orderby = `Date`;
    const searchListResult = await this._spservice.getPagedListItems(this.props.EmployeeList, _select, _filter, _orderby, _expand);
    destroyDataTable("staffTable")
    destroyDataTable("staffTable1")
    destroyDataTable("teamTimesheet-export")
    if (this.state.isPending) {
      var data = searchListResult?.filter((item) => item.Status.toLowerCase() === "waiting for approval")
      data.reverse();
      let filterdata = [];
      data.forEach(element => {
        if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
          filterdata.push(element)
        }
      });
      this.setState({ TeamListData: filterdata ? filterdata : [] });
      this.callDataTable("staffTable", this.state.groupColumn);
      this.calcTotalEfforts();
      this.calcEfforts();
    }
    else {
      searchListResult.reverse()
      let filterdata = [];
      searchListResult.forEach(element => {
        if (moment(element.Date).format("YYYY-MM-DD") >= this.state.fromDate && moment(element.Date).format("YYYY-MM-DD") <= this.state.toDate) {
          filterdata.push(element)
        }

      });
      this.setState({ TeamListData: filterdata ? filterdata : [] })
      this.callDataTable("staffTable1", this.state.groupColumn);
      this.calcTotalAllEfforts();
      this.calcEfforts();
    }
  }
  private handleClear = async () => {
    $('#projectName2').val(null).trigger("change");
    $('#projectName3').val(null).trigger("change");
    this.setState({ project: [], dateRange: 0, errorMessageFromDate: "", errorMessageToDate: "", PeopleId: 0, statusFilter: "", groupColumn: "" });
    await this.loadDefaultDates("0");
    this.state.isPending ? this.getPendingTimesheets() : this.getAllTimesheets()
  }
  //Actions
  private getItemlist = async (index) => {
    document.getElementById("comments1").classList.remove("border-danger");

    var employeeListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items(${index})?$select=*,Resource/Title&$expand=Resource`
    const employeeListResult = await this._spservice.get(employeeListURL);
    $("#newDate1").val(`${this._spservice.moment(employeeListResult.data.Date).format("YYYY-MM-DD")}`);
    var listURL = `${this._spservice.rootSite}${this._spservice.rootURL}/_api/web/lists/getbytitle('${this.props.ProjectList}')/items(${employeeListResult.data.ProjectID})?$select=*,Client/Title&$expand=Client`;
    const listResult = await this._spservice.get(listURL);
    if (listResult.ok) {
      let clientname = listResult.data.ClientId ? '[' + listResult.data.Client.Title + '] ' : '';
      $("#projectName1").val(clientname + listResult.data.Title);
    }
    //$("#description1").val(employeeListResult.data.TaskDescription);
    ($('#description1') as any).summernote('disable');
    ($('#description1') as any).summernote('code', employeeListResult.data.TaskDescription);
    $("#efforts1").val(employeeListResult.data.Effort);
    $("#taskType1").val(employeeListResult.data.TaskTypeId);
    $("#update").val(employeeListResult.data.Id);
    $("#comments1").val(employeeListResult.data.Comments);
    $("h4#myCenterModalLabel").text(employeeListResult.data.Resource.Title ? employeeListResult.data.Resource.Title : "Employee Name");
    this.setState({ status: employeeListResult.data.Status ? employeeListResult.data.Status.toLowerCase() : "" })
  }
  private handleApprove = async () => {
    var metadata = {
      Comments: $("#comments1").val(),
      Status: "Approved",
      Resubmit: false
    };
    var postURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items(${$("#update").val()})`
    await this._spservice.update(postURL, metadata);
    $('#custom-modal1').modal('hide');
    this.state.isPending ? this.getPendingTimesheets() : this.getAllTimesheets()
    this.setState({
      successMessage: "Approved Successfully"
    })
    // , () => {
    //   $('#myAlert').addClass('show');
    //   $('#myAlert').show();
    // })
    $("comments1").val("");
  }
  private handleReject = async () => {
    let validateFields: IValidationField[] = [
      { Id: "comments1", Type: "text" },
    ];
    const validationResult = this._spservice.utils.validateFields(validateFields, true);
    if (!validationResult.IsValid) {
      return
    }
    var metadata = {
      Comments: $("#comments1").val(),
      Status: "Rejected",
      Resubmit: false
    };
    var postURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items(${$("#update").val()})`
    await this._spservice.update(postURL, metadata);
    $('#custom-modal1').modal('hide');
    this.state.isPending ? this.getPendingTimesheets() : this.getAllTimesheets()
    this.setState({
      successMessage: "Rejected Successfully"
    })
    // , () => {
    //   $('#myAlert').addClass('show');
    //   $('#myAlert').show();
    // })
    $("comments1").val("");
  }
  private handleRevise = async () => {
    let validateFields: IValidationField[] = [
      { Id: "comments1", Type: "text" },
    ];
    const validationResult = this._spservice.utils.validateFields(validateFields, true);
    if (!validationResult.IsValid) {
      return
    }
    var metadata = {
      Comments: $("#comments1").val(),
      Status: "Revised",
      Resubmit: false
    };
    var postURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items(${$("#update").val()})`
    await this._spservice.update(postURL, metadata);
    $('#custom-modal1').modal('hide');
    this.state.isPending ? this.getPendingTimesheets() : this.getAllTimesheets()
    this.setState({
      successMessage: "Revised Successfully"
    })
    // , () => {
    //   $('#myAlert').addClass('show');
    //   $('#myAlert').show();
    // })
    $("comments1").val("");
  }
  private async calendar() {
    $(".hide-show-search").hide();
    $('#Pendingshow').hide();
    $('#Allshow').hide();
    $('#myTabContent').hide();
    $("#calendar").show();
    $(".fc-button").css("border", "rgb(34, 110, 183)");
    $(".fc-button").css("background-color", "rgb(34, 110, 183)");
    var arr = [];
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    var filter = "";
      this.state.staff.forEach((ele, index) => {
        if (filter == "") {
          if (index == 0)
            filter = `ResourceId eq ${ele.Id}`;
          else
            filter += ` or ResourceId eq ${ele.Id}`
        } else {
          if (index == 0)
            filter += ` and ResourceId eq ${ele.Id}`;
          else
            filter += ` or ResourceId eq ${ele.Id}`
        }
      });
    // var teamListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items?$select=*,Project/Title,Resource/Title,TaskType/Title&$expand=Project,Resource,TaskType&$filter=${filter}`
    var _select = [{ fieldType: null, key: '*,Resource/Title,TaskType/Title,FieldValuesAsText/TaskDescription' }];
    var _expand = ["Resource,TaskType,FieldValuesAsText"];
    var _orderby = `Date`;
    var _filter = `${filter}`;
    const teamListResult1 = await this._spservice.getPagedListItems(this.props.EmployeeList, _select, _filter, _orderby, _expand);
    var pendingData = teamListResult1;
    pendingData.map((item) => {
      let Project;
      this.state.ProjectList.map(proitem => {
        if (proitem.Id == item.ProjectID) {
          Project = proitem.Title;
        }
      })
      let Description = item.TaskDescription;
      arr.push({
        title: `${Project}: ${item.Effort}hours`,
        name: item.Resource.Title,
        date: moment(item.Date).format("YYYY-MM-DD"),
        effort: item.Effort,
        task: item.TaskType.Title,
        project: Project,
        extendedProps: {
          description: Description
        }
      })
    })
    let leavefilter = "";
    console.log(this.state.staff)
      this.state.staff.forEach((ele, index) => {
        if (leavefilter == "") {
          if (index == 0)
            leavefilter = `AuthorId eq ${ele.Id}`;
          else
            leavefilter += ` or AuthorId eq ${ele.Id}`
        } else {
          if (index == 0)
            leavefilter += ` and AuthorId eq ${ele.Id}`;
          else
            leavefilter += ` or AuthorId eq ${ele.Id}`
        }
      });
    console.log(leavefilter)
    // var teamListURL = `${this._spservice.absoluteUrl}/_api/web/lists/getbytitle('${this.props.EmployeeList}')/items?$select=*,Project/Title,Resource/Title,TaskType/Title&$expand=Project,Resource,TaskType&$filter=${filter}`
    // let User = this.props.context.pageContext.legacyPageContext.userId;
    let url = `${this._spservice.rootSite}${this._spservice.rootURL}/LMS/_api/web/lists/getbytitle('${this.props.Leaves}')/items?$select=*,Author/Title&$expand=Author&$filter=${leavefilter}`;
    let leaveValue = await this._spservice.get(url);
    let data = leaveValue.data.value;
    console.log(data);
    data?.map((item) => {
      let leave, color;
      if (item.Status === "Approved") {
        leave = 'Leave - ' + item.Author.Title
        color = "#c2c6cc"
      } else {
        leave = 'Waiting for Approval - ' + item.Author.Title
        color = "#ff0000"
      }
      // date: `${moment(item.StartDate).format("YYYY-MM-DD").toString()}'-'${moment(item.EndDate).format("YYYY-MM-DD").toString()}`,
      var startDay = new Date(item.StartDate);
      var endDay = new Date(item.EndDate);
      while (startDay <= endDay) {
        arr.push({
          title: leave,
          name: item.Author.Title,
          date: moment(startDay).format("YYYY-MM-DD"),
          overlap: true,
          color: color,
          extendedProps: {
            description: "leave"
          }
        })
        var newDate = startDay.setDate(startDay.getDate() + 1);
        startDay = new Date(newDate);
      }
    })
    this.setState({ event: arr }, () => {
      console.log(this.state.event);
    })
  }
  handleMouseEnter = (info) => {
    console.log(info.event.extendedProps.description);
    tooltipInstance = info.event.extendedProps.description;
    let Name = info.event.extendedProps.name;
    let task = info.event.extendedProps.task;
    let effort = info.event.extendedProps.effort;
    let Project = info.event.extendedProps.project;
    if (info.event.extendedProps.description) {
      $(info.el).tooltip({
        html: true,
        title: tooltipInstance == "leave" ? '<div style="text-align: left">' + "Name: " + Name + '</div>' : '<div style="text-align: left">' + "Name: " + Name + '</br>' + "Project: " + Project + '</br>' + "Task Type: " + task + '</br>' + "Effort: " + effort + "hours" + '</br>' + "Description: " + tooltipInstance + '</div>',
        placement: "left",
        trigger: "hover",
        container: "body"
      });
      $(info.el).tooltip('show');
    }
  };
  //calculation  
  private calcTotalEfforts() {
    let total = 0;
    // var efforts = document.getElementsByClassName("pendingefforts");
    // for (var i = 0; i < efforts.length; i++){	 
    //   total = +total + (+efforts[i].innerHTML);
    // }
    let efforts = this.state.TeamListData;
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
    let efforts = this.state.TeamListData;
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
    let efforts = this.state.TeamListData;
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

  public render(): React.ReactElement<ITeamTimesheetProps> {
    const { TeamListData, ProjectList, TaskType, status, successMessage, alertMessage } = this.state
    return (
      <>
        <div id="loader"></div>
        <div className="content-page">
          <div className="content">
            <div className="container-fluid pl-0 pr-0"></div>
            <div className="row mb-2 add-bar" style={{ paddingBottom: "19px" }}>
              <div className="col-sm-12 col-md-12 float-left ml-4">
                <span style={{
                  color: "#000000", right: "10px",
                  top: "58px", fontSize: "20px", cursor: "pointer"
                }} onClick={(e) => this.handleHamBurger(0)}> &#9776; </span>
                <span className='pl-3' style={{ fontSize: "20px", fontWeight: "bold" }}>My Team</span>
                <button onClick={this.handleExcel} type="button" className="btn btn-primary waves-effect waves-light float-right" style={{ backgroundColor: "#226EB7", marginRight: "52px", minHeight: "42px", minWidth: "110px" }}>
                  <i className="fa fa-download mr-2" />Export</button>
              </div>
            </div>

            <div className="row">
              <div className='row'>
                <ul className="nav nav-tabs" id="myTab" role="tablist" style={{ margin: "15px 0px 15px 39px", width: "96%" }}>
                  <li className="nav-item" role="presentation">
                    <button className="nav-link" id="pending-tab" data-toggle="tab" data-target="#pending" type="button" role="tab" aria-controls="pending" aria-selected="true" onClick={() => { this.getPendingTimesheets(); this.ppl.state.selectedPersons = [] }}>Pending</button>
                  </li>
                  <li className="nav-item" role="presentation">
                    <button className="nav-link active" id="all-tab" data-toggle="tab" data-target="#all" type="button" role="tab" aria-controls="all" aria-selected="false" onClick={() => { this.getAllTimesheets(); this.ppl.state.selectedPersons = [] }}>All</button>
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
                <div className="tab-pane fade show" id="pending" role="tabpanel" aria-labelledby="pending-tab">
                  <div className="row" id="search-filter" style={{ display: "flex", margin: "10px", marginTop: "30px", paddingRight: "35px" }}>
                    <div className="col-sm-3 col-md-3" >
                      <label htmlFor="position" id="label">Team Members </label>
                      {/* <div className="form-control"> */}
                      <PeoplePicker
                        context={this.props.context}
                        personSelectionLimit={1}
                        groupName={""} // Leave this blank in case you want to filter from all users
                        showtooltip={false}
                        required={false}
                        ensureUser={true}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        onChange={this.getPeople}
                        resolveDelay={1000}
                        styles={this.pickerStylesSingle}
                        ref={c => (this.ppl = c)} />
                      {/* </div> */}
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
                      <label id="label">From Date </label>
                      <input className='form-control datepicker place-holder' type="date" id="fromDate" name="fromDate"
                        value={this.state.fromDate} onChange={this.handleFromDateChange} max={moment().format("YYYY-MM-DD")}></input>
                      <span style={{ color: "#D8000C" }}>{this.state.errorMessageFromDate || null}</span>
                    </div>
                    <div className="col-sm-3 col-md-3">
                      <label id="label">To Date </label>
                      <input className='form-control datepicker place-holder' type="date" id="toDate" name="toDate"
                        value={this.state.toDate} onChange={this.handleToDateChange} max={moment().format("YYYY-MM-DD")}></input>
                      <span style={{ color: "#D8000C" }}>{this.state.errorMessageToDate || null}</span>
                    </div>
                    <div className="col-sm-3 col-md-3 mt-4">
                      <label htmlFor="position" id="label">Project</label>
                      <select className="form-control place-holder multiselect" id="projectName2" multiple={true}>
                        {ProjectList && ProjectList.map((item, index) => (
                          <option value={item.Id}>{`${item.ClientId ? '[' + item.Client.Title + ']' : ''} ${item.Title}`}</option>
                        ))}
                      </select>
                    </div>
                    <div className="col-sm-3 col-md-3 mt-4">
                      <label id="label">Group By</label>
                      <select className="form-control place-holder" id="example-select"
                        value={this.state.groupColumn} onChange={this.handleGroupBy}>
                        <option value="" selected>None</option>
                        <option value="0">Resource</option>
                        <option value="1">Project</option>
                        <option value="3">Task Type</option>
                      </select>
                    </div>
                    <div className="col-sm-3 col-md-3 mt-4"></div>
                    <div className="col-xs-12 col-sm-3 mt-2">
                      <div className="row col-6 p-0 m-0 float-right"><div className="col-xs-12 col-sm-12" style={{ padding: "0px 0px 0px 15px" }}><button type="button" className="btn waves-effect waves-light mt-3" style={{ width: "100%", height: "96%", color: "#226EB7", border: "1px solid" }} onClick={() => { this.ppl.state.selectedPersons = []; this.handleClear() }}><i className="fa fa-undo mr-1" /> Reset</button>
                        {/* <div className='row'>
                        <div className="col-xs-12 col-sm-6" style={{ padding: "0px 15px 0px 8px" }}>
                          <button type="button" style={{ width: "100%", height: "100%", backgroundColor: "#226EB7" }} className="btn btn-primary waves-effect waves-light mt-3" data-dismiss="modal"
                            onClick={this.handleSearch}><i className="fa fa-search mr-1" /> Search</button></div>
                        <div className="col-xs-12 col-sm-6" style={{ padding: "0px 0px 0px 15px" }}>
                          <button type="button" style={{ width: "100%", height: "96%", color: "#226EB7", border: "1px solid" }} className="btn waves-effect waves-light mt-3"
                            onClick={() => { this.ppl.state.selectedPersons = []; this.handleClear() }}><i className="fa fa-undo mr-1" /> Reset</button>*/}
                      </div>
                      </div>
                    </div>
                  </div>
                  <div className="row" id='Pendingshow'>
                    <div className="col-xl-12">
                      <div className="card-box bg-trans">
                        <div className="row pt-1">
                          <div className="col-xs-6 col-md-6 pl-4">
                            <span className="pl-4" style={{ fontWeight: "800", color: "#958F8F", fontSize: "13px" }}>{TeamListData.length ? `${TeamListData.length} Items Found` : "0 Items Found"}</span>
                            {/* <span className="pl-4" style={{ color: "#958F8F" }}>{EmployeeList.length ? `${EmployeeList.length} Items` : "0 Items"}</span> */}
                          </div>
                        </div>
                        <div className="table-responsive pt-3" style={{ cursor: "pointer", padding: "0px 26px 0px 30px" }}>
                          <table id="staffTable" className="table table-hover product_table">
                            <thead>
                              <tr className="table-active">
                                <th>Resource</th>
                                <th>Project</th>
                                <th>Task Description</th>
                                <th>Task Type</th>
                                <th>Effort (in hours)</th>
                                <th>Date</th>
                              </tr>
                            </thead>
                            <tbody>
                              {TeamListData && TeamListData.map((item, index) => (
                                <tr data-toggle="modal" data-backdrop="static" data-keyboard="false"
                                  onClick={() => this.getItemlist(item.Id)} data-target="#custom-modal1">
                                  <td className="">{item.Resource ? item.Resource.Title : ""}</td>
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
                                </tr>
                              ))}
                            </tbody>
                            <tfoot>
                              <tr>
                                <td></td>
                                <td></td>
                                <td style={{ fontWeight: "bold" }}>{TeamListData.length ? "Total Efforts: " : ""}</td>
                                <td></td>
                                <td style={{ fontWeight: "bold" }}>{TeamListData.length ? <span id="totalpendingefforts"></span> : null}</td>
                                <td></td>
                              </tr>
                            </tfoot>
                          </table>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="tab-pane fade show active" id="all" role="tabpanel" aria-labelledby="all-tab">
                  <div className="row" id="search-filter1" style={{ display: "flex", margin: "10px", marginTop: "30px", paddingRight: "35px" }}>
                    <div className="col-sm-3 col-md-3" >
                      <label htmlFor="position" id="label">Team Members </label>
                      {/* <div className="form-control"> */}
                      <PeoplePicker
                        context={this.props.context}
                        personSelectionLimit={1}
                        groupName={""} // Leave this blank in case you want to filter from all users
                        showtooltip={false}
                        required={false}
                        ensureUser={true}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        onChange={this.getPeople}
                        resolveDelay={1000}
                        styles={this.pickerStylesSingle}
                        ref={c => (this.ppl = c)} />
                      {/* </div> */}
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
                      <label id="label">From Date </label>
                      <input className='form-control datepicker place-holder' type="date" id="fromDate" name="fromDate"
                        value={this.state.fromDate} onChange={this.handleFromDateChange} max={moment().format("YYYY-MM-DD")}></input>
                      <span style={{ color: "#D8000C" }}>{this.state.errorMessageFromDate || null}</span>
                    </div>
                    <div className="col-sm-3 col-md-3">
                      <label id="label">To Date </label>
                      <input className='form-control datepicker place-holder' type="date" id="toDate" name="toDate"
                        value={this.state.toDate} onChange={this.handleToDateChange} max={moment().format("YYYY-MM-DD")}></input>
                      <span style={{ color: "#D8000C" }}>{this.state.errorMessageToDate || null}</span>
                    </div>
                    <div className="col-sm-3 col-md-3 mt-4">
                      <label htmlFor="position" id="label">Project</label>
                      <select className="form-control place-holder multiselect" id="projectName3" multiple={true}>
                        {ProjectList && ProjectList.map((item, index) => (
                          <option value={item.Id}>{`${item.ClientId ? '[' + item.Client.Title + ']' : ''} ${item.Title}`}</option>
                        ))}
                      </select>
                    </div>
                    <div className="col-sm-3 col-md-3 mt-4">
                      <label htmlFor="position" id="label">Status </label>
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
                        <option value="0">Resource</option>
                        <option value="1">Project</option>
                        <option value="3">Task Type</option>
                        <option value="5">Status</option>
                      </select>
                    </div>
                    <div className="col-xs-12 col-sm-3 mt-2">
                      {/*  <div className='row'>
                   <div className="col-xs-12 col-sm-6" style={{ padding: "0px 15px 0px 8px" }}>
                          <button type="button" style={{ width: "100%", height: "100%", backgroundColor: "#226EB7" }} className="btn btn-primary waves-effect waves-light mt-3" data-dismiss="modal"
                            onClick={this.handleSearch}><i className="fa fa-search mr-1" /> Search</button></div>
                     */}
                      <div className="row col-6 p-0 m-0 float-right"><div className="col-xs-12 col-sm-12" style={{ padding: "0px 0px 0px 15px" }}><button type="button" className="btn waves-effect waves-light mt-3" style={{ width: "100%", height: "96%", color: "#226EB7", border: "1px solid" }} onClick={() => { this.ppl.state.selectedPersons = []; this.handleClear() }}><i className="fa fa-undo mr-1" /> Reset</button>
                        {/* <div className="col-xs-12 col-sm-6" style={{ padding: "0px 0px 0px 15px" }}>
                          <button type="button" style={{ width: "100%", height: "96%", color: "#226EB7", border: "1px solid" }} className="btn waves-effect waves-light mt-3"
                            onClick={() => { this.ppl.state.selectedPersons = []; this.handleClear() }}><i className="fa fa-undo mr-1" /> Reset</button>
                    </div>*/}</div>
                      </div>
                    </div>
                  </div>
                  <div className="row" id='Allshow'>
                    <div className="col-xl-12">
                      <div className="card-box bg-trans">
                        <div className="row pt-1">
                          <div className="col-xs-6 col-md-6 pl-4">
                            <span className="pl-4" style={{ fontWeight: "800", color: "#958F8F", fontSize: "13px" }}>{TeamListData.length ? `${TeamListData.length} Items Found` : "0 Items Found"}</span>
                          </div>
                        </div>
                        <div className="table-responsive pt-3" style={{ cursor: "pointer", padding: "0px 26px 0px 30px" }}>
                          <table id="staffTable1" className="table table-hover product_table">
                            <thead>
                              <tr className="table-active">
                                <th>Resource</th>
                                <th>Project</th>
                                <th>Task Description</th>
                                <th>Task Type</th>
                                <th>Effort (in hours)</th>
                                <th>Status</th>
                                <th>Date</th>
                              </tr>
                            </thead>
                            <tbody>
                              {TeamListData && TeamListData.map((item, index) => (
                                <tr data-toggle="modal" data-backdrop="static" data-keyboard="false"
                                  onClick={() => this.getItemlist(item.Id)} data-target="#custom-modal1">
                                  <td className="">{item.Resource ? item.Resource.Title : ""}</td>
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
                                <td></td>
                                <td style={{ fontWeight: "bold" }}>{TeamListData.length ? "Total Efforts : " : ""}</td>
                                <td></td>
                                <td style={{ fontWeight: "bold" }}>{TeamListData.length ? <span id="totalallefforts"></span> : null}</td>
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

        <div className="modal left fade" id="custom-modal1" tabIndex={-1} role="dialog" aria-hidden="true">
          <div className="modal-dialog modal-dialog-centered">
            <div className="modal-content">
              <div className="modal-header bg-light" style={{ display: "block" }}>
                <div className="row">
                  <div className="col-1 col-md-1 col-sm-1" style={{ padding: "3px" }}>
                    <span className="close" data-dismiss="modal" aria-hidden="true" style={{ textAlign: "right", fontSize: "24px !important" }}>×</span>
                  </div>
                  <div className="col-10 col-md-10 col-sm-10">
                    <h4 className="modal-title" id="myCenterModalLabel"></h4>
                  </div>
                </div>
              </div>
              <div className="modal-body p-4">
                <form id="myform1">
                  <div className="form-group">
                    <label htmlFor="name">Date</label>
                    <input className='form-control datepicker' disabled={true} type="date" id="newDate1" name="newDate1"
                    ></input>
                  </div>
                  <div className="form-group">
                    <label htmlFor="name">Project</label>
                    <input type="text" className="form-control" id="projectName1" disabled={true} onChange={this.handleSearch} />
                  </div>
                  <div className="form-group" id='summernotedescription'>
                    <label htmlFor="position">Task Description</label>
                    <div className="summernote" id='description1'></div>
                    {/* <textarea className="form-control" id="description1" name="description1" disabled={true} rows={4} /> */}
                  </div>
                  <div className="form-group">
                    <label htmlFor="category">Effort (in hours)</label>
                    <input type="number" className="form-control" min="1" max="12" id="efforts1" name="efforts1" disabled={true} />
                  </div>
                  <div className="form-group">
                    <label htmlFor="category">Task Type</label>
                    <select className="form-control" id="taskType1" disabled={true}>
                      <option value={0}>-- Select --</option>
                      {TaskType && TaskType.map((item, index) => (
                        <option value={item.Id}>{item.Title}</option>
                      ))}
                    </select>
                  </div>
                  <div className="form-group">
                    <label htmlFor="position">Comments</label>
                    <textarea className="form-control" id="comments1" name="comments1" rows={4}
                      disabled={status === "waiting for approval" ? false : true} />
                  </div>
                </form>
              </div>
              <div className="modal-footer text-center display-block p-4" id="update">
                {status === "waiting for approval" ? <>
                  <button type="button" className='btn btn-danger' style={{ width: "31%" }} onClick={this.handleReject} >Reject</button>
                  <button type="button" className="btn btn-secondary" style={{ width: "31%" }} onClick={this.handleRevise} >Revise</button>
                  <button type="button" className='btn btn-success' style={{ width: "32%" }} onClick={this.handleApprove} >Approve</button>
                </>
                  :
                  <button type="button" className="btn btn-secondary" data-dismiss="modal" style={{ width: "100%" }}>Close</button>}

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

        <table id="teamTimesheet-export" className="display nowrap d-none" style={{ width: "100%" }}>
          <thead>
            <tr>
              <th>Date</th>
              <th>Resource</th>
              <th>Project</th>
              <th>Task Description</th>
              <th>Task Type</th>
              <th>Effort (in hours)</th>
              <th>{this.state.isPending ? "" : "Status"}</th>
            </tr>
          </thead>
          <tbody>
            {TeamListData && TeamListData.map((item, i) => {
              return (
                <tr key={i}>
                  <td>{item.Date ? this._spservice.moment(item.Date).format("YYYY-MM-DD") : ""}</td>
                  <td>{item.Resource ? item.Resource.Title : ""}</td>
                  <td>
                    {ProjectList && ProjectList.map(proitem => {
                      if (proitem.Id == item.ProjectID) {
                        return proitem.ClientId ? ('[' + proitem.Client.Title + '] ' + proitem.Title) : proitem.Title;
                      }
                    })}
                  </td>
                  <td dangerouslySetInnerHTML={{ __html: item.TaskDescription }}></td>
                  <td className="">{item.TaskType ? item.TaskType.Title : ""}</td>
                  <td className="exportefforts">{item.Effort ? item.Effort : 0}</td>
                  <td>{this.state.isPending ? "" : item.Status}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr>
              <td></td>
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
