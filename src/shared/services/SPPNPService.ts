import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, IRegionalSettings, ISiteUser, PermissionKind } from '@pnp/sp/presets/all';
import SPPNPLogger from '../services/SPPNPLogger';
import { IUserPermissions } from '../models/IUserPermissions';
import { ListTemplateType } from '../utils/enums';
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { IEventData } from "../models/IEventData";
import * as moment from 'moment';
import parseRecurrentEvent from "./parseRecurrentEvent";


export default class SPPNPService {

  public _logger: SPPNPLogger;
  public sp: typeof sp; 

    constructor(private context: WebPartContext) {
        sp.setup({
            spfxContext: this.context
        });
        this._logger = new SPPNPLogger();
        this.sp = sp;
    }

    public async addComment(){
      sp.web.lists.getByTitle("distest").items.get();
      const item = await sp.web.getFileByServerRelativeUrl("/Lists/distest/Parent Item").getItem();

      // as an example, or any of the below options
      await item.comments();
    }

    public async getListItems(selectedList: string, selectedFields: any[]) {
        try {
            let selectQuery: any[] = ['Id'];
            let expandQuery: any[] = [];
            let listItems = [];
            let items: any;
            for (var i = 0; i < selectedFields.length; i++) {
                switch (selectedFields[i].fieldType) {
                    case 'SP.FieldUser':
                        selectQuery.push(`${selectedFields[i].key}/Title,${selectedFields[i].key}/EMail,${selectedFields[i].key}/Name`);
                        expandQuery.push(selectedFields[i].key);
                        break;
                    case 'SP.FieldLookup':
                        selectQuery.push(`${selectedFields[i].key}/Title`);
                        expandQuery.push(selectedFields[i].key);
                        break;
                    case 'SP.Field':
                        selectQuery.push('Attachments,AttachmentFiles');
                        expandQuery.push('AttachmentFiles');
                        break;
                    default:
                        selectQuery.push(selectedFields[i].key);
                        break;
                }
            }
            items = await sp.web.lists.getById(selectedList).items
                .select(selectQuery.join())
                .expand(expandQuery.join())
                .top(4999)
                .getPaged();
            listItems = items.results;
            while (items.hasNext) {
                items = await items.getNext();
                listItems = [...listItems, ...items.results];
            }
            return listItems;
        } catch (err) {
            Promise.reject(err);
        }
    }

    public async getFields(selectedList: string): Promise<any> {
        try {
            const allFields: any[] = await sp.web.lists
                .getById(selectedList)
                .fields
                .filter("Hidden eq false and ReadOnlyField eq false and Title ne 'Content Type' and Title ne 'Attachments'")
                .get();
            return allFields;
        }
        catch (err) {
            Promise.reject(err);
        }
    }

    public async getUserProfileUrl(loginName: string) {
        try {
            const properties = await sp.profiles.getPropertiesFor(loginName);
            const profileUrl = properties['PictureUrl'];
            return profileUrl;
        }
        catch (err) {
            Promise.reject(err);
        }
    }

    /**
   *
   * @private
   * @returns {Promise<string>}
   * @memberof spservices
   */
  public async getLocalTime(date: string | Date): Promise<string> {
    try {
      const localTime = await sp.web.regionalSettings.timeZone.utcToLocalTime(date);
      return localTime;
    }
    catch (error) {
      return Promise.reject(error);
    }
  }

  /**
   *
   * @private
   * @returns {Promise<string>}
   * @memberof spservices
   */
  public async getUtcTime(date: string | Date): Promise<string> {
    try {
      const utcTime = await sp.web.regionalSettings.timeZone.localTimeToUTC(date);
      return utcTime;
    }
    catch (error) {
      return Promise.reject(error);
    }
  }

  /**
   *
   * @param {number} userId
   * @param {string} siteUrl
   * @returns {Promise<SiteUser>}
   * @memberof spservices
   */
  public async getUserById(userId: number, siteUrl: string): Promise<ISiteUserInfo> {
    let results: ISiteUserInfo = null;

    if (!userId && !siteUrl) {
      return null;
    }

    try {
      //const web = new Web(siteUrl);
      results = await sp.web.siteUsers.getById(userId).get();
      //results = await web.siteUsers.getByLoginName(userId).get();
    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }

  /**
   *
   *
   * @param {string} loginName
   * @param {string} siteUrl
   * @returns {Promise<SiteUser>}
   * @memberof spservices
   */
  public async getUserByLoginName(loginName: string, siteUrl: string): Promise<ISiteUserInfo> {
    let results: ISiteUserInfo = null;

    if (!loginName && !siteUrl) {
      return null;
    }

    try {
      //const web = new Web(siteUrl);
      await sp.web.ensureUser(loginName);
      results = await sp.web.siteUsers.getByLoginName(loginName).get();
      //results = await web.siteUsers.getByLoginName(userId).get();
    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }
  /**
   *
   * @param {string} loginName
   * @returns
   * @memberof spservices
   */
  public async getUserProfilePictureUrl(loginName: string) {
    let results: any = null;
    try {
      results = await sp.profiles.usingCaching().getPropertiesFor(loginName);
    } catch (error) {
      results = null;
    }
    return results.PictureUrl;
  }

  /**
   *
   * @param {string} listName
   * @returns {Promise<IUserPermissions>}
   * @memberof spservices
   */
  public async getUserPermissions(listName: string): Promise<IUserPermissions> {
    let hasPermissionAdd: boolean = false;
    let hasPermissionEdit: boolean = false;
    let hasPermissionDelete: boolean = false;
    let hasPermissionView: boolean = false;
    let userPermissions: IUserPermissions = undefined;
    try {
      //const web = new Web(siteUrl);
      const userEffectivePermissions = await sp.web.lists.getByTitle(listName).effectiveBasePermissions.get();
      // ...
      hasPermissionAdd = sp.web.lists.getByTitle(listName).hasPermissions(userEffectivePermissions, PermissionKind.AddListItems);
      hasPermissionDelete = sp.web.lists.getByTitle(listName).hasPermissions(userEffectivePermissions, PermissionKind.DeleteListItems);
      hasPermissionEdit = sp.web.lists.getByTitle(listName).hasPermissions(userEffectivePermissions, PermissionKind.EditListItems);
      hasPermissionView = sp.web.lists.getByTitle(listName).hasPermissions(userEffectivePermissions, PermissionKind.ViewListItems);
      userPermissions = { hasPermissionAdd: hasPermissionAdd, hasPermissionEdit: hasPermissionEdit, hasPermissionDelete: hasPermissionDelete, hasPermissionView: hasPermissionView };

    } catch (error) {
      return Promise.reject(error);
    }
    return userPermissions;
  }
  /**
   *
   * @param {string} siteUrl
   * @returns
   * @memberof spservices
   */
  public async getSiteLists(siteUrl: string) {

    let results: any[] = [];

    // if (!siteUrl) {
    //   return [];
    // }

    try {
      //const web = new Web(siteUrl);
      results = await sp.web.lists.select("Title", "ID").filter('BaseTemplate eq ' + ListTemplateType.CustomList).get();
    } catch (error) {
      return Promise.reject(error);
    }
    return results;
  }

  /**
   *
   * @private
   * @param {string} siteUrl
   * @returns
   * @memberof spservices
   */
  public async getSiteRegionalSettingsTimeZone(siteUrl: string) {
    let regionalSettings: IRegionalSettings;
    try {
      ////const web = new Web(siteUrl);
      //regionalSettings = await sp.web.regionalSettings.timeZone.usingCaching().get();

    } catch (error) {
      return Promise.reject(error);
    }
    return regionalSettings;
  }
  
  /**
   * ensure that a list exists. If it doesn't it will be created with the provided title (the rest of the settings will be default):
   * @public
   * @param {string} listName
   * @returns
   * @memberof spservices
   */
  public async ensureList(listName: string) {

    // ensure that a list exists. If it doesn't it will be created with the provided title (the rest of the settings will be default):
    const listEnsureResult = await sp.web.lists.ensure(listName, listName, ListTemplateType.CustomList);

    // check if the list was created, or if it already existed:
    if (listEnsureResult.created) {
        this._logger.info(listName+ " was created!");
    } else {
        this._logger.info(listName+ " already existed!");
    }

    // work on the created/updated list
    const r = await listEnsureResult.list.select("Id")();

    // return the Id
    return r.Id;
  }

  /**
   *
   * @param {string} siteUrl
   * @param {string} listName
   * @param {string} fieldInternalName
   * @returns {Promise<{ key: string, text: string }[]>}
   * @memberof spservices
   */
   public async getChoiceFieldOptions(siteUrl: string, listName: string, fieldInternalName: string): Promise<{ key: string, text: string }[]> {
    let fieldOptions: { key: string, text: string }[] = [];
    try {
      //const web = new Web(siteUrl);
      const results: any = await sp.web.lists.getByTitle(listName)
        .fields
        .getByInternalNameOrTitle(fieldInternalName)
        .select("Title", "InternalName", "Choices")
        .get();
      if (results && results.Choices.length > 0) {
        for (const option of results.Choices) {
          fieldOptions.push({
            key: option,
            text: option
          });
        }
      }
    } catch (error) {
      return Promise.reject(error);
    }
    return fieldOptions;
  }

  /**
   *
   * @private
   * @returns
   * @memberof spservices
   */
   public async colorGenerate() {

    var hexValues = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e"];
    var newColor = "#";

    for (var i = 0; i < 6; i++) {
      var x = Math.round(Math.random() * 14);

      var y = hexValues[x];
      newColor += y;
    }
    return newColor;
  }

  /**
   *
   * @param {string} siteUrl
   * @param {string} listName
   * @param {Date} eventStartDate
   * @param {Date} eventEndDate
   * @returns {Promise< IEventData[]>}
   * @memberof spservices
   */
   public async getEvents(listName: string, eventStartDate: Date | null, eventEndDate: Date | null, top: number = 0, filter = {}): Promise<IEventData[]> {

    let events: IEventData[] = [];
    try {

      // Get Category Field Choices
      const categoryDropdownOption = await this.getChoiceFieldOptions("", listName, 'Category');
      let categoryColor: { category: string, color: string }[] = [];
      for (const cat of categoryDropdownOption) {
        categoryColor.push({ category: cat.text, color: await this.colorGenerate() });
      }

      var today = moment().format();
      var DatesInUtc = false;
      //const web = new Web(siteUrl);
      var query = 
        `<View>
          <ViewFields>
            <FieldRef Name='RecurrenceData'/>
            <FieldRef Name='Duration'/>
            <FieldRef Name='Author'/>
            <FieldRef Name='Category'/>
            <FieldRef Name='Description'/>
            <FieldRef Name='ParticipantsPicker'/>
            <FieldRef Name='Geolocation'/>
            <FieldRef Name='ID'/>
            <FieldRef Name='EndDate'/>
            <FieldRef Name='EventDate'/>
            <FieldRef Name='Id'/>
            <FieldRef Name='Location'/>
            <FieldRef Name='Title'/>
            <FieldRef Name='fAllDayEvent'/>
            <FieldRef Name='EventType'/>
            <FieldRef Name='UID' />
            <FieldRef Name='fRecurrence' />
          </ViewFields>
          <Query>
            <Where>
              <Or>
                <DateRangesOverlap>
                    <FieldRef Name='EventDate' />
                    <FieldRef Name='EndDate' />
                    <FieldRef Name='RecurrenceID' />
                    <Value Type='DateTime'>
                      <Today />
                    </Value>
                </DateRangesOverlap>
                <And>
                 <Geq>
                   <FieldRef Name='EventDate' />
                   <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventStartDate).format('YYYY-MM-DD')}</Value>
                 </Geq>
                 <Leq>
                   <FieldRef Name='EventDate' />
                   <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventEndDate).format('YYYY-MM-DD')}</Value>
                 </Leq>
                </And>
              </Or>
            </Where>
            <OrderBy>
              <FieldRef Name='EventDate' />
            </OrderBy>
          </Query>
          <QueryOptions>
            <CalendarDate>${today}</CalendarDate>
            <ExpandRecurrence>TRUE</ExpandRecurrence>
            <RecurrenceOrderBy>TRUE</RecurrenceOrderBy>
            <ViewAttributes Scope='RecursiveAll'/>
          </QueryOptions>
        </View>`;

      const results = await sp.web.lists.getByTitle(listName).usingCaching().renderListDataAsStream(
        {
          //DatesInUtc: true,
          ViewXml: query,
          // ViewXml: `<View><ViewFields><FieldRef Name='RecurrenceData'/><FieldRef Name='Duration'/><FieldRef Name='Author'/><FieldRef Name='Category'/><FieldRef Name='Description'/><FieldRef Name='ParticipantsPicker'/><FieldRef Name='Geolocation'/><FieldRef Name='ID'/><FieldRef Name='EndDate'/><FieldRef Name='EventDate'/><FieldRef Name='ID'/><FieldRef Name='Location'/><FieldRef Name='Title'/><FieldRef Name='fAllDayEvent'/><FieldRef Name='EventType'/><FieldRef Name='UID' /><FieldRef Name='fRecurrence' /></ViewFields>
          // <Query>
          // <Where>
          //   <And>
          //     <Geq>
          //       <FieldRef Name='EventDate' />
          //       <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventStartDate).format('YYYY-MM-DD')}</Value>
          //     </Geq>
          //     <Leq>
          //       <FieldRef Name='EventDate' />
          //       <Value IncludeTimeValue='false' Type='DateTime'>${moment(eventEndDate).format('YYYY-MM-DD')}</Value>
          //     </Leq>
          //     </And>
          // </Where>
          // </Query>
          // <RowLimit Paged=\"FALSE\">2000</RowLimit>
          // </View>`
          //OverrideViewXml: "<QueryOptions><ExpandRecurrence>TRUE</ExpandRecurrence></QueryOptions>"
        }
      );

      console.log(results);

      if (results && results.Row.length > 0) {
        let row = results.Row;
        if(Object.keys(filter).length > 0){
          row = row.filter(f => {
            let isMatching = true;
            Object.keys(filter).forEach(function(key,index) {
              // key: the name of the object key
              // index: the ordinal position of the key within the object
              let oValue = f[key];
              let fValue = filter[key];
              let valid = (oValue.toLowerCase().indexOf(fValue.toLowerCase()) > -1) ? true : false;
              if(!valid)
                isMatching = false;
            });
            return isMatching;
          });
        }

        let event: any = '';
        for (event of row) {
          const eventDate = DatesInUtc ? await this.getLocalTime(event.EventDate) : event.EventDate;
          const endDate = DatesInUtc ? await this.getLocalTime(event.EndDate) : event.EndDate;
          const initialsArray: string[] = event.Author[0].title.split(' ');
          const initials: string = initialsArray[0].charAt(0) + initialsArray[initialsArray.length - 1].charAt(0);
          const userPictureUrl = await this.getUserProfilePictureUrl(`i:0#.f|membership|${event.Author[0].email}`);
          const attendees: number[] = [];
          const first: number = event.Geolocation != undefined ? event.Geolocation.indexOf('(') + 1 : 0;
          const last: number = event.Geolocation != undefined ? event.Geolocation.indexOf(')') : 0;
          const geo = event.Geolocation != undefined ? event.Geolocation.substring(first, last) : '0 0';
          const geolocation = geo.split(' ');
          const CategoryColorValue: any[] = categoryColor.filter((value) => {
            return value.category == event.Category;
          });
          const isAllDayEvent: boolean = event["fAllDayEvent.value"] === "1";

          for (const attendee of event.ParticipantsPicker) {
            attendees.push(parseInt(attendee.id));
          }

          events.push({
            Id: event.ID,
            ID: event.ID,
            EventType: event.EventType,
            title: await this.deCodeHtmlEntities(event.Title),
            Description: event.Description,
            EventDate: isAllDayEvent ? new Date(event.EventDate.slice(0, -1)) : new Date(eventDate),
            EndDate: isAllDayEvent ? new Date(event.EndDate.slice(0, -1)) : new Date(endDate),
            location: event.Location,
            ownerEmail: event.Author[0].email,
            ownerPhoto: userPictureUrl ?
              `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${event.Author[0].email}&UA=0&size=HR96x96` : '',
            ownerInitial: initials,
            color: CategoryColorValue.length > 0 ? CategoryColorValue[0].color : '#1a75ff', // blue default
            ownerName: event.Author[0].title,
            attendes: attendees,
            fAllDayEvent: isAllDayEvent,
            geolocation: { Longitude: parseFloat(geolocation[0]), Latitude: parseFloat(geolocation[1]) },
            Category: event.Category,
            Duration: event.Duration,
            RecurrenceData: event.RecurrenceData ? await this.deCodeHtmlEntities(event.RecurrenceData) : "",
            fRecurrence: event.fRecurrence,
            RecurrenceID: event.RecurrenceID ? event.RecurrenceID : undefined,
            MasterSeriesItemID: event.MasterSeriesItemID,
            UID: event.UID.replace("{", "").replace("}", ""),
            monthName : moment(event.EventDate).format("MMMM"),
            monthShortName : moment(event.EventDate).format("MMM"),
            eventDay : moment(event.EventDate).format("DD"),
          });
        }
        // const mapEvents = async () : Promise<boolean> => {
        //     for (event of results.Row) {
        //       const eventDate = await this.getLocalTime(event.EventDate);
        //       const endDate = await this.getLocalTime(event.EndDate);
        //       const initialsArray: string[] = event.Author[0].title.split(' ');
        //       const initials: string = initialsArray[0].charAt(0) + initialsArray[initialsArray.length - 1].charAt(0);
        //       const userPictureUrl = await this.getUserProfilePictureUrl(`i:0#.f|membership|${event.Author[0].email}`);
        //       const attendees: number[] = [];
        //       const first: number = event.Geolocation.indexOf('(') + 1;
        //       const last: number = event.Geolocation.indexOf(')');
        //       const geo = event.Geolocation.substring(first, last);
        //       const geolocation = geo.split(' ');
        //       const CategoryColorValue: any[] = categoryColor.filter((value) => {
        //         return value.category == event.Category;
        //       });
        //       const isAllDayEvent: boolean = event["fAllDayEvent.value"] === "1";

        //       for (const attendee of event.ParticipantsPicker) {
        //         attendees.push(parseInt(attendee.id));
        //       }

        //       events.push({
        //         Id: event.ID,
        //         ID: event.ID,
        //         EventType: event.EventType,
        //         title: await this.deCodeHtmlEntities(event.Title),
        //         Description: event.Description,
        //         EventDate: isAllDayEvent ? new Date(event.EventDate.slice(0, -1)) : new Date(eventDate),
        //         EndDate: isAllDayEvent ? new Date(event.EndDate.slice(0, -1)) : new Date(endDate),
        //         location: event.Location,
        //         ownerEmail: event.Author[0].email,
        //         ownerPhoto: userPictureUrl ?
        //           `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${event.Author[0].email}&UA=0&size=HR96x96` : '',
        //         ownerInitial: initials,
        //         color: CategoryColorValue.length > 0 ? CategoryColorValue[0].color : '#1a75ff', // blue default
        //         ownerName: event.Author[0].title,
        //         attendes: attendees,
        //         fAllDayEvent: isAllDayEvent,
        //         geolocation: { Longitude: parseFloat(geolocation[0]), Latitude: parseFloat(geolocation[1]) },
        //         Category: event.Category,
        //         Duration: event.Duration,
        //         RecurrenceData: event.RecurrenceData ? await this.deCodeHtmlEntities(event.RecurrenceData) : "",
        //         fRecurrence: event.fRecurrence,
        //         RecurrenceID: event.RecurrenceID ? event.RecurrenceID : undefined,
        //         MasterSeriesItemID: event.MasterSeriesItemID,
        //         UID: event.UID.replace("{", "").replace("}", ""),
        //       });
        //     }
        //   return true;
        // };
        //Checks to see if there are any results saved in local storage
        // if(window.localStorage.getItem("eventResult")){
        //   //if there is a local version - compares it to the current version 
        //   if(window.localStorage.getItem("eventResult") === JSON.stringify(results)){
        //     //No update needed use current savedEvents
        //     events = JSON.parse(window.localStorage.getItem("calendarEventsWithLocalTime"));
        //   }else{
        //     //update local storage
        //     window.localStorage.setItem("eventResult", JSON.stringify(results));
        //     //when they are not equal then we loop through the results and maps them to IEventData
        //     /* tslint:disable:no-unused-expression */
        //     await mapEvents() ? window.localStorage.setItem("calendarEventsWithLocalTime", JSON.stringify(events)) : null;           
        //   }
        // }else{
        //   //if there is no local storage of the events we create them
        //   window.localStorage.setItem("eventResult", JSON.stringify(results));
        //   //we also needs to map through the events the first time and save the mapped version to local storage
        //   await mapEvents() ? window.localStorage.setItem("calendarEventsWithLocalTime", JSON.stringify(events)) : null;           
        // }
      }
      let parseEvt: parseRecurrentEvent = new parseRecurrentEvent();
      events = parseEvt.parseEvents(events, eventStartDate, null, top);
       
      //Return Data
      events.sort(function (a, b) {
        var aDate = moment(a.EventDate);
        var bDate = moment(b.EventDate);
        return aDate.isAfter(bDate) ? 1 : (aDate.isBefore(bDate) ? -1 : 0);
      });
      
      if(top > 0 && events.length > 0) 
        return events.slice(0, top);
      else
        return events;
    } catch (error) {
      console.dir(error);
      return Promise.reject(error);
    }
  }

  /**
   *
   *
   * @param {string} siteUrl
   * @param {string} listName
   * @param {number} eventId
   * @returns {Promise<IEventData>}
   * @memberof spservices
   */
   public async getEvent(listName: string, eventId: number, siteUrl: string = ""): Promise<IEventData> {
    let returnEvent: IEventData = undefined;
    try {
      //const web = new Web(siteUrl);

      //"Title","fRecurrence", "fAllDayEvent","EventDate", "EndDate", "Description","ID", "Location","Geolocation","ParticipantsPickerId"
      const event = await sp.web.lists.getByTitle(listName).items.getById(eventId)
        .select("RecurrenceID", "MasterSeriesItemID", "Id", "ID", "ParticipantsPickerId", "EventType", "Title", "Description", "EventDate", "EndDate", "Location", "Author/SipAddress", "Author/Title", "fAllDayEvent", "fRecurrence", "RecurrenceData", "RecurrenceData", "Duration", "Category", "UID", "RegistrationCloseDate", "SendEmail", "ContactPersonId", "Scope", "ContactPerson/Title", "ContactPerson/Id")
        .expand("Author,ContactPerson")
        .get();

      const eventDate = await this.getLocalTime(event.EventDate);
      const endDate = await this.getLocalTime(event.EndDate);

      returnEvent = {
        Id: event.ID,
        ID: event.ID,
        EventType: event.EventType,
        title: await this.deCodeHtmlEntities(event.Title),
        Description: event.Description ? event.Description : '',
        EventDate: new Date(eventDate),
        EndDate: new Date(endDate),
        location: event.Location,
        ownerEmail: event.Author.SipAddress,
        ownerPhoto: "",
        ownerInitial: '',
        color: '',
        ownerName: event.Author.Title,
        attendes: event.ParticipantsPickerId,
        fAllDayEvent: event.fAllDayEvent,
        geolocation: { Longitude: event.Geolocation ? event.Geolocation.Longitude : 0, Latitude: event.Geolocation ? event.Geolocation.Latitude : 0 },
        Category: event.Category,
        Duration: event.Duration,
        UID: event.UID,
        RecurrenceData: event.RecurrenceData ? await this.deCodeHtmlEntities(event.RecurrenceData) : "",
        fRecurrence: event.fRecurrence,
        RecurrenceID: event.RecurrenceID,
        MasterSeriesItemID: event.MasterSeriesItemID,
        RegistrationCloseDate: event.RegistrationCloseDate != null ? await this.getLocalTime(event.RegistrationCloseDate) : null,
        Scope: event.Scope,
        SendEmail: event.SendEmail,
        ContactPersonId: event.ContactPersonId != null ? event.ContactPersonId : null,
        ContactPerson: event.ContactPerson != null ? event.ContactPerson : null
      };
    } 
    catch (error) {
      return Promise.reject(error);
    }
    return returnEvent;
  }

  public async deCodeHtmlEntities(string: string) {

    const HtmlEntitiesMap = {
      "'": "&#39;",
      "<": "&lt;",
      ">": "&gt;",
      " ": "&nbsp;",
      "¡": "&iexcl;",
      "¢": "&cent;",
      "£": "&pound;",
      "¤": "&curren;",
      "¥": "&yen;",
      "¦": "&brvbar;",
      "§": "&sect;",
      "¨": "&uml;",
      "©": "&copy;",
      "ª": "&ordf;",
      "«": "&laquo;",
      "¬": "&not;",
      "®": "&reg;",
      "¯": "&macr;",
      "°": "&deg;",
      "±": "&plusmn;",
      "²": "&sup2;",
      "³": "&sup3;",
      "´": "&acute;",
      "µ": "&micro;",
      "¶": "&para;",
      "·": "&middot;",
      "¸": "&cedil;",
      "¹": "&sup1;",
      "º": "&ordm;",
      "»": "&raquo;",
      "¼": "&frac14;",
      "½": "&frac12;",
      "¾": "&frac34;",
      "¿": "&iquest;",
      "À": "&Agrave;",
      "Á": "&Aacute;",
      "Â": "&Acirc;",
      "Ã": "&Atilde;",
      "Ä": "&Auml;",
      "Å": "&Aring;",
      "Æ": "&AElig;",
      "Ç": "&Ccedil;",
      "È": "&Egrave;",
      "É": "&Eacute;",
      "Ê": "&Ecirc;",
      "Ë": "&Euml;",
      "Ì": "&Igrave;",
      "Í": "&Iacute;",
      "Î": "&Icirc;",
      "Ï": "&Iuml;",
      "Ð": "&ETH;",
      "Ñ": "&Ntilde;",
      "Ò": "&Ograve;",
      "Ó": "&Oacute;",
      "Ô": "&Ocirc;",
      "Õ": "&Otilde;",
      "Ö": "&Ouml;",
      "×": "&times;",
      "Ø": "&Oslash;",
      "Ù": "&Ugrave;",
      "Ú": "&Uacute;",
      "Û": "&Ucirc;",
      "Ü": "&Uuml;",
      "Ý": "&Yacute;",
      "Þ": "&THORN;",
      "ß": "&szlig;",
      "à": "&agrave;",
      "á": "&aacute;",
      "â": "&acirc;",
      "ã": "&atilde;",
      "ä": "&auml;",
      "å": "&aring;",
      "æ": "&aelig;",
      "ç": "&ccedil;",
      "è": "&egrave;",
      "é": "&eacute;",
      "ê": "&ecirc;",
      "ë": "&euml;",
      "ì": "&igrave;",
      "í": "&iacute;",
      "î": "&icirc;",
      "ï": "&iuml;",
      "ð": "&eth;",
      "ñ": "&ntilde;",
      "ò": "&ograve;",
      "ó": "&oacute;",
      "ô": "&ocirc;",
      "õ": "&otilde;",
      "ö": "&ouml;",
      "÷": "&divide;",
      "ø": "&oslash;",
      "ù": "&ugrave;",
      "ú": "&uacute;",
      "û": "&ucirc;",
      "ü": "&uuml;",
      "ý": "&yacute;",
      "þ": "&thorn;",
      "ÿ": "&yuml;",
      "Œ": "&OElig;",
      "œ": "&oelig;",
      "Š": "&Scaron;",
      "š": "&scaron;",
      "Ÿ": "&Yuml;",
      "ƒ": "&fnof;",
      "ˆ": "&circ;",
      "˜": "&tilde;",
      "Α": "&Alpha;",
      "Β": "&Beta;",
      "Γ": "&Gamma;",
      "Δ": "&Delta;",
      "Ε": "&Epsilon;",
      "Ζ": "&Zeta;",
      "Η": "&Eta;",
      "Θ": "&Theta;",
      "Ι": "&Iota;",
      "Κ": "&Kappa;",
      "Λ": "&Lambda;",
      "Μ": "&Mu;",
      "Ν": "&Nu;",
      "Ξ": "&Xi;",
      "Ο": "&Omicron;",
      "Π": "&Pi;",
      "Ρ": "&Rho;",
      "Σ": "&Sigma;",
      "Τ": "&Tau;",
      "Υ": "&Upsilon;",
      "Φ": "&Phi;",
      "Χ": "&Chi;",
      "Ψ": "&Psi;",
      "Ω": "&Omega;",
      "α": "&alpha;",
      "β": "&beta;",
      "γ": "&gamma;",
      "δ": "&delta;",
      "ε": "&epsilon;",
      "ζ": "&zeta;",
      "η": "&eta;",
      "θ": "&theta;",
      "ι": "&iota;",
      "κ": "&kappa;",
      "λ": "&lambda;",
      "μ": "&mu;",
      "ν": "&nu;",
      "ξ": "&xi;",
      "ο": "&omicron;",
      "π": "&pi;",
      "ρ": "&rho;",
      "ς": "&sigmaf;",
      "σ": "&sigma;",
      "τ": "&tau;",
      "υ": "&upsilon;",
      "φ": "&phi;",
      "χ": "&chi;",
      "ψ": "&psi;",
      "ω": "&omega;",
      "ϑ": "&thetasym;",
      "ϒ": "&Upsih;",
      "ϖ": "&piv;",
      "–": "&ndash;",
      "—": "&mdash;",
      "‘": "&lsquo;",
      "’": "&rsquo;",
      "‚": "&sbquo;",
      "“": "&ldquo;",
      "”": "&rdquo;",
      "„": "&bdquo;",
      "†": "&dagger;",
      "‡": "&Dagger;",
      "•": "&bull;",
      "…": "&hellip;",
      "‰": "&permil;",
      "′": "&prime;",
      "″": "&Prime;",
      "‹": "&lsaquo;",
      "›": "&rsaquo;",
      "‾": "&oline;",
      "⁄": "&frasl;",
      "€": "&euro;",
      "ℑ": "&image;",
      "℘": "&weierp;",
      "ℜ": "&real;",
      "™": "&trade;",
      "ℵ": "&alefsym;",
      "←": "&larr;",
      "↑": "&uarr;",
      "→": "&rarr;",
      "↓": "&darr;",
      "↔": "&harr;",
      "↵": "&crarr;",
      "⇐": "&lArr;",
      "⇑": "&UArr;",
      "⇒": "&rArr;",
      "⇓": "&dArr;",
      "⇔": "&hArr;",
      "∀": "&forall;",
      "∂": "&part;",
      "∃": "&exist;",
      "∅": "&empty;",
      "∇": "&nabla;",
      "∈": "&isin;",
      "∉": "&notin;",
      "∋": "&ni;",
      "∏": "&prod;",
      "∑": "&sum;",
      "−": "&minus;",
      "∗": "&lowast;",
      "√": "&radic;",
      "∝": "&prop;",
      "∞": "&infin;",
      "∠": "&ang;",
      "∧": "&and;",
      "∨": "&or;",
      "∩": "&cap;",
      "∪": "&cup;",
      "∫": "&int;",
      "∴": "&there4;",
      "∼": "&sim;",
      "≅": "&cong;",
      "≈": "&asymp;",
      "≠": "&ne;",
      "≡": "&equiv;",
      "≤": "&le;",
      "≥": "&ge;",
      "⊂": "&sub;",
      "⊃": "&sup;",
      "⊄": "&nsub;",
      "⊆": "&sube;",
      "⊇": "&supe;",
      "⊕": "&oplus;",
      "⊗": "&otimes;",
      "⊥": "&perp;",
      "⋅": "&sdot;",
      "⌈": "&lceil;",
      "⌉": "&rceil;",
      "⌊": "&lfloor;",
      "⌋": "&rfloor;",
      "⟨": "&lang;",
      "⟩": "&rang;",
      "◊": "&loz;",
      "♠": "&spades;",
      "♣": "&clubs;",
      "♥": "&hearts;",
      "♦": "&diams;"
    };

    var entityMap = HtmlEntitiesMap;
    for (var key in entityMap) {
      var entity = entityMap[key];
      var regex = new RegExp(entity, 'g');
      string = string.replace(regex, key);
    }
    string = string.replace(/&quot;/g, '"');
    string = string.replace(/&amp;/g, '&');
    return string;
  }
}
