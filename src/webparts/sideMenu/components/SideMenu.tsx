import styles from './SideMenu.module.scss';
import { ISideMenuProps } from './ISideMenuProps';
import SPService from '../../../shared/services/SPService';
import { MSGraphClient } from '@microsoft/sp-http';
import * as React from 'react';

export interface ISideMenuState {
  teamMembers: any[];
  profile: any;
  isAdmin: boolean;
  isHR: boolean;
  isManager: boolean;
  userMail:any;
}

export default class SideMenu extends React.Component<ISideMenuProps, ISideMenuState> {
  private _spservice: SPService;
  constructor(props) {
    super(props);
    this._spservice = new SPService(this.props.context);
    this.state = {
      teamMembers: [],
      profile: {},
      isAdmin: false,
      isHR: false,
      isManager: false,
      userMail:this.props.context.pageContext.legacyPageContext.userEmail,
    }
  }

  componentDidMount = () => {
    this.getProfile();
    this.checkPermission();
    //this.getTeamList();
  }

  componentDidUpdate = () => {
    var path = window.location.href;
    $('ul a').each(function () {
      if (this.href === path) {
        var $parent = $(this).parent();
        $parent.addClass('active');
      }
    })
  }

  // private async GetSiteAdmin() {
  //   var siteURL = `${this._spservice.absoluteUrl}/_api/web/currentUser/issiteadmin`
  //   const siteResult = await this._spservice.get(siteURL);
  //   this.setState({ isAdmin: siteResult.data.value })
  // }

  // private getTeamList = async () => {
  //   let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
  //   var teamListURL = `${this._spservice.rootSite}${this._spservice.rootURL}/_api/web/lists/getbytitle('${this.props.TeamList}')/items?$select=Staff/Title,Staff/Id&$expand=Staff&$filter=ManagerId eq ${currentUserId}`
  //   const teamListResult = await this._spservice.get(teamListURL);
  //   console.log(teamListResult.data.value)
  //   this.setState({ teamMembers: teamListResult.data.ok ?  teamListResult.data.value : [] })
  // }

  private getProfile = async () => {
    var profileURL = `${this._spservice.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`
    const profileResult = await this._spservice.get(profileURL);
    this.setState({ profile: profileResult.data })
  }

  private async checkPermission() {
    let currentUserId = parseInt(this.props.context.pageContext.legacyPageContext.userId);
    var adminURL = `${this._spservice.absoluteUrl}/_api/web/GetUserById('${currentUserId}')/Groups?$filter=Title eq '${this.props.adminGroup}'`;
    const adminResult = await this._spservice.get(adminURL);
    var hrURL = `${this._spservice.absoluteUrl}/_api/web/GetUserById('${currentUserId}')/Groups?$filter=Title eq '${this.props.hrGroup}'`;
    const hrResult = await this._spservice.get(hrURL);
    this.getManager();
    this.setState({
      isAdmin: (adminResult.ok && adminResult.data.value.length > 0) ? true : false,
      isHR: (hrResult.ok && hrResult.data.value.length > 0) ? true : false,
    });
  }
  private async getManager(): Promise<any> {
    new Promise<any>((resolve, reject) => {
      this.props.context.msGraphClientFactory.getClient()
        .then((client: MSGraphClient) => {
          // https://graph.microsoft.com/v1.0/me/manager
          // let UserEmail = this.props.context.pageContext.legacyPageContext.userEmail;
          client.api('users').version('beta').filter(`mail eq '${this.state.userMail}'`).get((error, response: any, rawResponse?: any) => {
            let managetId=response.value[0].id
            console.log(response)
            this.getStaff(managetId);
            if (error) {
              // reject(error);
              resolve(null);
              return
            }
            else {
              // resolve(response.value[0].manager.displayName)
              console.log(response.value);
              let currentUser =this.props.context.pageContext.legacyPageContext.userEmail;
              response.value?.map((element) => {
                let userManager = element?.manager?.mail
                if (userManager === currentUser) {
                  console.log(element?.manager?.mail);
                  this.setState({ isManager: true })

                }
              })
            }
          });
        });
    });
  }
  private async getStaff(managetId): Promise<any> {
    new Promise<any>((resolve, reject) => {
      this.props.context.msGraphClientFactory.getClient()
        .then(async (client: MSGraphClient) => {
          // https://graph.microsoft.com/v1.0/me/manager
              client.api(`users/${managetId}/directReports`).version('beta').get((error, response: any, rawResponse?: any) => {
                console.log(response)
            if (error) {
              // reject(error);.filter(`manager.mail eq '${UserEmail}'`)
              resolve(null);
              return
            }
            else {
              // resolve(response.value[0].manager.displayName)
              console.log(response.value);
              if(response.value.length !== 0){
                this.setState({ isManager: true })
              }
            }
          });
        });
    });
  }

  private handleHamBurger = (id) => {
    if (id === 0) {
      document.getElementById("left-side-menu").style.width = "0";
      document.getElementById("profile").style.width = "0";
      document.getElementById("image").style.width = "0";
      document.getElementById("profilephoto").style.width = "0px";
      document.getElementById("left-side-menu-mini").style.width = "84px";
      document.getElementById("profile-mini").style.width = "78px";
      document.getElementById("image-mini").style.width = "80px";
      document.getElementById("profilephoto-mini").style.width = "76px";
      document.getElementById("side-menu-mini").style.display = "";
      document.getElementById("Qt-logo-mini").style.display = "";
      $('ul#side-menu li.active').css("width", "0");
      $('#mini-sidebar').css("margin-left", "85px");
      // $('.left-side-menu').hide();
      // $('.left-side-menu-mini').show();
    }
    else if (id === 1) {
      document.getElementById("left-side-menu-mini").style.width = "0";
      document.getElementById("profile-mini").style.width = "0";
      document.getElementById("image-mini").style.width = "0";
      document.getElementById("profilephoto-mini").style.width = "0px";
      document.getElementById("left-side-menu").style.width = "240px";
      document.getElementById("profile").style.width = "225px";
      document.getElementById("image").style.width = "238px";
      document.getElementById("profilephoto").style.width = "78px";
      document.getElementById("side-menu-mini").style.display = "none";
      document.getElementById("Qt-logo-mini").style.display = "none";
      $('ul#side-menu li.active').css("width", "225px");
      $('#mini-sidebar').css("margin-left", "240px");
      // $('.left-side-menu-mini').hide();
      // $('.left-side-menu').show();
    }
  }

  public render(): React.ReactElement<ISideMenuProps> {
    const { profile, isAdmin, isHR, isManager } = this.state;
    return (
      <>
        <div className="left-side-menu" id="left-side-menu" style={{}}>
          <div className="h-75" data-simplebar>
            <div id="sidebar-menu">
              <div className="auth-logo mb-3 text-center">
                <a href={this._spservice.rootURL}>
                  <img src={this._spservice.absoluteUrl + "/SiteAssets/images/Qantler-logo.svg"} alt="" />
                </a>
                {/* <span style={{
                  color: "#ffffff", position: "absolute", right: "10px",
                  top: "58px", fontSize: "20px", cursor: "pointer"
                }} onClick={(e) => this.handleHamBurger(0)}> &#9776; </span> */}
              </div>
              <div style={{ borderBottom: "1px dashed #4371A9" }}></div>
              <div className='profile' id="profile">
                <div className='row' style={{ height: "100%" }}>
                  <div className='col-sm-4 col-md-4'>
                    <img id="profilephoto" src={profile.PictureUrl ? `${this._spservice.absoluteUrl}/_layouts/15/userphoto.aspx?accountname=${profile.Email}&amp;size=M` : `${this._spservice.absoluteUrl}/SiteAssets/images/Default-image.png`} style={{
                      width: "70px",
                      height: "70px",
                      margin: "13px 0px 0px 13px",
                      borderRadius: "50px"
                    }}></img></div>
                  <div className='col-sm-8 col-md-8' style={{
                    margin: "auto",
                    textOverflow: "ellipsis",
                    overflow: "hidden",
                    color: "#ffffff",
                    paddingLeft: "28px",
                  }}>
                    <span style={{ color: "FFFFFF" }}>{profile.DisplayName}</span><br></br>
                    <span style={{ color: "#dadada", fontSize: "12px" }}>{profile.Title}</span>
                  </div>
                </div>
              </div>
              <ul id="side-menu" style={{ listStyleType: "none", padding: "0" }}>
                <li className="nav-item" style={{ marginTop: "5px" }}>
                  <a className='' href={this._spservice.absoluteUrl + "/Pages/Timesheet.aspx"}>
                    <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="layer-group" className="svg-inline--fa fa-layer-group fa-w-16" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style={{
                      width: "30px",
                      height: "17px",
                      margin: "14px 0px 17px 20px",
                    }}>
                      <path fill="#dadada" d="M12.41 148.02l232.94 105.67c6.8 3.09 14.49 3.09 21.29 0l232.94-105.67c16.55-7.51 16.55-32.52 0-40.03L266.65 2.31a25.607 25.607 0 0 0-21.29 0L12.41 107.98c-16.55 7.51-16.55 32.53 0 40.04zm487.18 88.28l-58.09-26.33-161.64 73.27c-7.56 3.43-15.59 5.17-23.86 5.17s-16.29-1.74-23.86-5.17L70.51 209.97l-58.1 26.33c-16.55 7.5-16.55 32.5 0 40l232.94 105.59c6.8 3.08 14.49 3.08 21.29 0L499.59 276.3c16.55-7.5 16.55-32.5 0-40zm0 127.8l-57.87-26.23-161.86 73.37c-7.56 3.43-15.59 5.17-23.86 5.17s-16.29-1.74-23.86-5.17L70.29 337.87 12.41 364.1c-16.55 7.5-16.55 32.5 0 40l232.94 105.59c6.8 3.08 14.49 3.08 21.29 0L499.59 404.1c16.55-7.5 16.55-32.5 0-40z"></path>
                    </svg>
                    <span style={{ color: "#dadada" }}>My Timesheet</span>
                  </a>
                </li>
                {(isManager) ?
                  <li className="nav-item" style={{ marginTop: "5px" }}>
                    <a className='' href={this._spservice.absoluteUrl + "/Pages/MyTeam.aspx"}>
                      <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="user-friends" className="svg-inline--fa fa-user-friends fa-w-20" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 512" style={{
                        width: "30px",
                        height: "17px",
                        margin: "14px 0px 17px 20px",
                      }}>
                        <path fill="#dadada" d="M192 256c61.9 0 112-50.1 112-112S253.9 32 192 32 80 82.1 80 144s50.1 112 112 112zm76.8 32h-8.3c-20.8 10-43.9 16-68.5 16s-47.6-6-68.5-16h-8.3C51.6 288 0 339.6 0 403.2V432c0 26.5 21.5 48 48 48h288c26.5 0 48-21.5 48-48v-28.8c0-63.6-51.6-115.2-115.2-115.2zM480 256c53 0 96-43 96-96s-43-96-96-96-96 43-96 96 43 96 96 96zm48 32h-3.8c-13.9 4.8-28.6 8-44.2 8s-30.3-3.2-44.2-8H432c-20.4 0-39.2 5.9-55.7 15.4 24.4 26.3 39.7 61.2 39.7 99.8v38.4c0 2.2-.5 4.3-.6 6.4H592c26.5 0 48-21.5 48-48 0-61.9-50.1-112-112-112z"></path>
                      </svg>
                      <span style={{ color: "#dadada" }}>My Team</span>
                    </a>
                  </li>
                  : null}
                {/* {isAdmin ?
                  <li className="nav-item" style={{ marginTop: "5px" }}>
                    <a className='' href={`${this._spservice.rootSite}${this._spservice.rootURL}/Lists/Managers/AllItems.aspx`}>
                      <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="user-cog" className="svg-inline--fa fa-user-cog fa-w-20" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 512"
                        style={{
                          width: "30px",
                          height: "17px",
                          margin: "14px 0px 17px 20px",
                        }}>
                        <path fill="#dadada" d="M610.5 373.3c2.6-14.1 2.6-28.5 0-42.6l25.8-14.9c3-1.7 4.3-5.2 3.3-8.5-6.7-21.6-18.2-41.2-33.2-57.4-2.3-2.5-6-3.1-9-1.4l-25.8 14.9c-10.9-9.3-23.4-16.5-36.9-21.3v-29.8c0-3.4-2.4-6.4-5.7-7.1-22.3-5-45-4.8-66.2 0-3.3.7-5.7 3.7-5.7 7.1v29.8c-13.5 4.8-26 12-36.9 21.3l-25.8-14.9c-2.9-1.7-6.7-1.1-9 1.4-15 16.2-26.5 35.8-33.2 57.4-1 3.3.4 6.8 3.3 8.5l25.8 14.9c-2.6 14.1-2.6 28.5 0 42.6l-25.8 14.9c-3 1.7-4.3 5.2-3.3 8.5 6.7 21.6 18.2 41.1 33.2 57.4 2.3 2.5 6 3.1 9 1.4l25.8-14.9c10.9 9.3 23.4 16.5 36.9 21.3v29.8c0 3.4 2.4 6.4 5.7 7.1 22.3 5 45 4.8 66.2 0 3.3-.7 5.7-3.7 5.7-7.1v-29.8c13.5-4.8 26-12 36.9-21.3l25.8 14.9c2.9 1.7 6.7 1.1 9-1.4 15-16.2 26.5-35.8 33.2-57.4 1-3.3-.4-6.8-3.3-8.5l-25.8-14.9zM496 400.5c-26.8 0-48.5-21.8-48.5-48.5s21.8-48.5 48.5-48.5 48.5 21.8 48.5 48.5-21.7 48.5-48.5 48.5zM224 256c70.7 0 128-57.3 128-128S294.7 0 224 0 96 57.3 96 128s57.3 128 128 128zm201.2 226.5c-2.3-1.2-4.6-2.6-6.8-3.9l-7.9 4.6c-6 3.4-12.8 5.3-19.6 5.3-10.9 0-21.4-4.6-28.9-12.6-18.3-19.8-32.3-43.9-40.2-69.6-5.5-17.7 1.9-36.4 17.9-45.7l7.9-4.6c-.1-2.6-.1-5.2 0-7.8l-7.9-4.6c-16-9.2-23.4-28-17.9-45.7.9-2.9 2.2-5.8 3.2-8.7-3.8-.3-7.5-1.2-11.4-1.2h-16.7c-22.2 10.2-46.9 16-72.9 16s-50.6-5.8-72.9-16h-16.7C60.2 288 0 348.2 0 422.4V464c0 26.5 21.5 48 48 48h352c10.1 0 19.5-3.2 27.2-8.5-1.2-3.8-2-7.7-2-11.8v-9.2z"></path></svg>
                      <span style={{ color: "#dadada" }}>Managers</span>
                    </a>
                  </li> : null} */}
                {isAdmin ?
                  <li className="nav-item" style={{ marginTop: "5px" }}>
                    <a className='' href={this._spservice.absoluteUrl + "/Lists/TaskType/AllItems.aspx"}>
                      <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="user-cog" className="svg-inline--fa fa-user-cog fa-w-20" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 512"
                        style={{
                          width: "30px",
                          height: "17px",
                          margin: "14px 0px 17px 20px",
                        }}>
                        <path fill="#dadada" d="M610.5 373.3c2.6-14.1 2.6-28.5 0-42.6l25.8-14.9c3-1.7 4.3-5.2 3.3-8.5-6.7-21.6-18.2-41.2-33.2-57.4-2.3-2.5-6-3.1-9-1.4l-25.8 14.9c-10.9-9.3-23.4-16.5-36.9-21.3v-29.8c0-3.4-2.4-6.4-5.7-7.1-22.3-5-45-4.8-66.2 0-3.3.7-5.7 3.7-5.7 7.1v29.8c-13.5 4.8-26 12-36.9 21.3l-25.8-14.9c-2.9-1.7-6.7-1.1-9 1.4-15 16.2-26.5 35.8-33.2 57.4-1 3.3.4 6.8 3.3 8.5l25.8 14.9c-2.6 14.1-2.6 28.5 0 42.6l-25.8 14.9c-3 1.7-4.3 5.2-3.3 8.5 6.7 21.6 18.2 41.1 33.2 57.4 2.3 2.5 6 3.1 9 1.4l25.8-14.9c10.9 9.3 23.4 16.5 36.9 21.3v29.8c0 3.4 2.4 6.4 5.7 7.1 22.3 5 45 4.8 66.2 0 3.3-.7 5.7-3.7 5.7-7.1v-29.8c13.5-4.8 26-12 36.9-21.3l25.8 14.9c2.9 1.7 6.7 1.1 9-1.4 15-16.2 26.5-35.8 33.2-57.4 1-3.3-.4-6.8-3.3-8.5l-25.8-14.9zM496 400.5c-26.8 0-48.5-21.8-48.5-48.5s21.8-48.5 48.5-48.5 48.5 21.8 48.5 48.5-21.7 48.5-48.5 48.5zM224 256c70.7 0 128-57.3 128-128S294.7 0 224 0 96 57.3 96 128s57.3 128 128 128zm201.2 226.5c-2.3-1.2-4.6-2.6-6.8-3.9l-7.9 4.6c-6 3.4-12.8 5.3-19.6 5.3-10.9 0-21.4-4.6-28.9-12.6-18.3-19.8-32.3-43.9-40.2-69.6-5.5-17.7 1.9-36.4 17.9-45.7l7.9-4.6c-.1-2.6-.1-5.2 0-7.8l-7.9-4.6c-16-9.2-23.4-28-17.9-45.7.9-2.9 2.2-5.8 3.2-8.7-3.8-.3-7.5-1.2-11.4-1.2h-16.7c-22.2 10.2-46.9 16-72.9 16s-50.6-5.8-72.9-16h-16.7C60.2 288 0 348.2 0 422.4V464c0 26.5 21.5 48 48 48h352c10.1 0 19.5-3.2 27.2-8.5-1.2-3.8-2-7.7-2-11.8v-9.2z"></path></svg>
                      <span style={{ color: "#dadada" }}>Settings</span>
                    </a>
                  </li> : null}
              </ul>
            </div>
          </div>
          <div style={{ width: "100%", position: "absolute", bottom: "0" }}>
            <img style={{ transition: "0.3s" }} id="image" src={this._spservice.absoluteUrl + "/SiteAssets/images/Qantler-logo-footer.svg"} width={238} height={210} />
          </div>
        </div>


        <div className="left-side-menu-mini" id="left-side-menu-mini">
          <div className="h-75" data-simplebar>
            <div id="sidebar-menu">
              <div className="auth-logo mb-4 text-center" style={{ marginTop: "4px" }}>
                <img style={{ display: "none" }} src={this._spservice.absoluteUrl + "/SiteAssets/images/Qt-logo-mini.svg"} alt="" id="Qt-logo-mini" />
                {/* <span style={{
                  color: "#ffffff", position: "absolute", right: "5px",
                  top: "65px", fontSize: "20px", cursor: "pointer"
                }} onClick={(e) => this.handleHamBurger(1)}> &#9776; </span> */}
              </div>
              <div style={{ borderBottom: "1px dashed #4371A9" }}></div>
              <div className='profile-mini' id='profile-mini'>
                <div className='row'>
                  <div className='col-sm-4 col-md-4'>
                    <img id="profilephoto-mini" src={profile.PictureUrl ? `${this._spservice.absoluteUrl}/_layouts/15/userphoto.aspx?accountname=${profile.Email}&amp;size=M` : `${this._spservice.absoluteUrl}/SiteAssets/images/Default-image.png`} style={{
                      width: "0",
                      height: "70px",
                      margin: "15px 0px 0px 5px",
                      borderRadius: "50px"
                    }}></img>
                  </div>
                </div>
              </div>
              <ul id="side-menu-mini" style={{ listStyleType: "none", padding: "0px", display: "none" }}>
                <li className="nav-item" style={{ marginTop: "5px" }}>
                  <a className='' href={this._spservice.absoluteUrl + "/Pages/Timesheet.aspx"}>
                    <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="layer-group" className="svg-inline--fa fa-layer-group fa-w-16" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style={{
                      width: "30px",
                      height: "17px",
                      margin: "14px 0px 13px 22px",
                    }}>
                      <path fill="#dadada" d="M12.41 148.02l232.94 105.67c6.8 3.09 14.49 3.09 21.29 0l232.94-105.67c16.55-7.51 16.55-32.52 0-40.03L266.65 2.31a25.607 25.607 0 0 0-21.29 0L12.41 107.98c-16.55 7.51-16.55 32.53 0 40.04zm487.18 88.28l-58.09-26.33-161.64 73.27c-7.56 3.43-15.59 5.17-23.86 5.17s-16.29-1.74-23.86-5.17L70.51 209.97l-58.1 26.33c-16.55 7.5-16.55 32.5 0 40l232.94 105.59c6.8 3.08 14.49 3.08 21.29 0L499.59 276.3c16.55-7.5 16.55-32.5 0-40zm0 127.8l-57.87-26.23-161.86 73.37c-7.56 3.43-15.59 5.17-23.86 5.17s-16.29-1.74-23.86-5.17L70.29 337.87 12.41 364.1c-16.55 7.5-16.55 32.5 0 40l232.94 105.59c6.8 3.08 14.49 3.08 21.29 0L499.59 404.1c16.55-7.5 16.55-32.5 0-40z"></path>
                    </svg>
                  </a>
                </li>
                {(isManager || isAdmin || isHR) ?
                  <li className="nav-item" style={{ marginTop: "5px" }}>
                    <a className='' href={this._spservice.absoluteUrl + "/Pages/MyTeam.aspx"}>
                      <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="user-friends" className="svg-inline--fa fa-user-friends fa-w-20" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 512" style={{
                        width: "30px",
                        height: "17px",
                        margin: "14px 0px 13px 22px"
                      }}>
                        <path fill="#dadada" d="M192 256c61.9 0 112-50.1 112-112S253.9 32 192 32 80 82.1 80 144s50.1 112 112 112zm76.8 32h-8.3c-20.8 10-43.9 16-68.5 16s-47.6-6-68.5-16h-8.3C51.6 288 0 339.6 0 403.2V432c0 26.5 21.5 48 48 48h288c26.5 0 48-21.5 48-48v-28.8c0-63.6-51.6-115.2-115.2-115.2zM480 256c53 0 96-43 96-96s-43-96-96-96-96 43-96 96 43 96 96 96zm48 32h-3.8c-13.9 4.8-28.6 8-44.2 8s-30.3-3.2-44.2-8H432c-20.4 0-39.2 5.9-55.7 15.4 24.4 26.3 39.7 61.2 39.7 99.8v38.4c0 2.2-.5 4.3-.6 6.4H592c26.5 0 48-21.5 48-48 0-61.9-50.1-112-112-112z"></path>
                      </svg>
                    </a>
                  </li>
                  : null}
                {isAdmin ?
                  <li className="nav-item" style={{ marginTop: "5px" }}>
                    <a className='' href={`${this._spservice.rootSite}${this._spservice.rootURL}/Lists/Managers/AllItems.aspx`}>
                      <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="user-cog" className="svg-inline--fa fa-user-cog fa-w-20" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 512"
                        style={{
                          width: "30px",
                          height: "17px",
                          margin: "14px 0px 13px 22px"
                        }}>
                        <path fill="#dadada" d="M610.5 373.3c2.6-14.1 2.6-28.5 0-42.6l25.8-14.9c3-1.7 4.3-5.2 3.3-8.5-6.7-21.6-18.2-41.2-33.2-57.4-2.3-2.5-6-3.1-9-1.4l-25.8 14.9c-10.9-9.3-23.4-16.5-36.9-21.3v-29.8c0-3.4-2.4-6.4-5.7-7.1-22.3-5-45-4.8-66.2 0-3.3.7-5.7 3.7-5.7 7.1v29.8c-13.5 4.8-26 12-36.9 21.3l-25.8-14.9c-2.9-1.7-6.7-1.1-9 1.4-15 16.2-26.5 35.8-33.2 57.4-1 3.3.4 6.8 3.3 8.5l25.8 14.9c-2.6 14.1-2.6 28.5 0 42.6l-25.8 14.9c-3 1.7-4.3 5.2-3.3 8.5 6.7 21.6 18.2 41.1 33.2 57.4 2.3 2.5 6 3.1 9 1.4l25.8-14.9c10.9 9.3 23.4 16.5 36.9 21.3v29.8c0 3.4 2.4 6.4 5.7 7.1 22.3 5 45 4.8 66.2 0 3.3-.7 5.7-3.7 5.7-7.1v-29.8c13.5-4.8 26-12 36.9-21.3l25.8 14.9c2.9 1.7 6.7 1.1 9-1.4 15-16.2 26.5-35.8 33.2-57.4 1-3.3-.4-6.8-3.3-8.5l-25.8-14.9zM496 400.5c-26.8 0-48.5-21.8-48.5-48.5s21.8-48.5 48.5-48.5 48.5 21.8 48.5 48.5-21.7 48.5-48.5 48.5zM224 256c70.7 0 128-57.3 128-128S294.7 0 224 0 96 57.3 96 128s57.3 128 128 128zm201.2 226.5c-2.3-1.2-4.6-2.6-6.8-3.9l-7.9 4.6c-6 3.4-12.8 5.3-19.6 5.3-10.9 0-21.4-4.6-28.9-12.6-18.3-19.8-32.3-43.9-40.2-69.6-5.5-17.7 1.9-36.4 17.9-45.7l7.9-4.6c-.1-2.6-.1-5.2 0-7.8l-7.9-4.6c-16-9.2-23.4-28-17.9-45.7.9-2.9 2.2-5.8 3.2-8.7-3.8-.3-7.5-1.2-11.4-1.2h-16.7c-22.2 10.2-46.9 16-72.9 16s-50.6-5.8-72.9-16h-16.7C60.2 288 0 348.2 0 422.4V464c0 26.5 21.5 48 48 48h352c10.1 0 19.5-3.2 27.2-8.5-1.2-3.8-2-7.7-2-11.8v-9.2z"></path></svg>
                    </a>
                  </li> : null}
                {isAdmin ?
                  <li className="nav-item" style={{ marginTop: "5px" }}>
                    <a className='' href={this._spservice.absoluteUrl + "/Lists/TaskType/AllItems.aspx"}>
                      <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="user-cog" className="svg-inline--fa fa-user-cog fa-w-20" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 512"
                        style={{
                          width: "30px",
                          height: "17px",
                          margin: "14px 0px 13px 22px"
                        }}>
                        <path fill="#dadada" d="M610.5 373.3c2.6-14.1 2.6-28.5 0-42.6l25.8-14.9c3-1.7 4.3-5.2 3.3-8.5-6.7-21.6-18.2-41.2-33.2-57.4-2.3-2.5-6-3.1-9-1.4l-25.8 14.9c-10.9-9.3-23.4-16.5-36.9-21.3v-29.8c0-3.4-2.4-6.4-5.7-7.1-22.3-5-45-4.8-66.2 0-3.3.7-5.7 3.7-5.7 7.1v29.8c-13.5 4.8-26 12-36.9 21.3l-25.8-14.9c-2.9-1.7-6.7-1.1-9 1.4-15 16.2-26.5 35.8-33.2 57.4-1 3.3.4 6.8 3.3 8.5l25.8 14.9c-2.6 14.1-2.6 28.5 0 42.6l-25.8 14.9c-3 1.7-4.3 5.2-3.3 8.5 6.7 21.6 18.2 41.1 33.2 57.4 2.3 2.5 6 3.1 9 1.4l25.8-14.9c10.9 9.3 23.4 16.5 36.9 21.3v29.8c0 3.4 2.4 6.4 5.7 7.1 22.3 5 45 4.8 66.2 0 3.3-.7 5.7-3.7 5.7-7.1v-29.8c13.5-4.8 26-12 36.9-21.3l25.8 14.9c2.9 1.7 6.7 1.1 9-1.4 15-16.2 26.5-35.8 33.2-57.4 1-3.3-.4-6.8-3.3-8.5l-25.8-14.9zM496 400.5c-26.8 0-48.5-21.8-48.5-48.5s21.8-48.5 48.5-48.5 48.5 21.8 48.5 48.5-21.7 48.5-48.5 48.5zM224 256c70.7 0 128-57.3 128-128S294.7 0 224 0 96 57.3 96 128s57.3 128 128 128zm201.2 226.5c-2.3-1.2-4.6-2.6-6.8-3.9l-7.9 4.6c-6 3.4-12.8 5.3-19.6 5.3-10.9 0-21.4-4.6-28.9-12.6-18.3-19.8-32.3-43.9-40.2-69.6-5.5-17.7 1.9-36.4 17.9-45.7l7.9-4.6c-.1-2.6-.1-5.2 0-7.8l-7.9-4.6c-16-9.2-23.4-28-17.9-45.7.9-2.9 2.2-5.8 3.2-8.7-3.8-.3-7.5-1.2-11.4-1.2h-16.7c-22.2 10.2-46.9 16-72.9 16s-50.6-5.8-72.9-16h-16.7C60.2 288 0 348.2 0 422.4V464c0 26.5 21.5 48 48 48h352c10.1 0 19.5-3.2 27.2-8.5-1.2-3.8-2-7.7-2-11.8v-9.2z"></path></svg>
                    </a>
                  </li> : null}
              </ul>
            </div>
          </div>
          <div style={{ width: "100%", position: "absolute", bottom: "0" }}>
            <img id="image-mini" src={this._spservice.absoluteUrl + "/SiteAssets/images/Qt-footer-mini.svg"} width={0} height={80} />
          </div>
        </div>

      </>
    );
  }
}
