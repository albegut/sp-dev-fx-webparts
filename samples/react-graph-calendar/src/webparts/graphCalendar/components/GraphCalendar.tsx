import { EventInput } from '@fullcalendar/core';
import allLocales from '@fullcalendar/core/locales-all';
import dayGridPlugin from '@fullcalendar/daygrid';
import FullCalendar from '@fullcalendar/react';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as strings from "GraphCalendarWebPartStrings";
import { extendMoment } from 'moment-range';
import * as moment from 'moment-timezone';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import styles from './GraphCalendar.module.scss';
import { IGraphCalendarProps } from './IGraphCalendarProps';
import * as microsoftTeams from '@microsoft/teams-js';

const { range } = extendMoment(moment);

interface IGraphCalendarState {
  events: EventInput[];
  height: number;
  currentActiveStartDate: Date | undefined;
  currentActiveEndDate: Date | undefined;
  isEventDetailsOpen: boolean;
  currentSelectedEvent: EventInput | undefined;
  groupId: string | undefined;
  tabType: TabType;
}

enum TabType {
  TeamsTab,
  PersonalTab
}

export default class GraphCalendar extends React.Component<IGraphCalendarProps, IGraphCalendarState> {
  private calendar = React.createRef<FullCalendar>();
  private calendarLang: string = '';

  /**
   * Builds the Component with the provided properties
   * @param props Properties of the web part
   */
  constructor(props: IGraphCalendarProps) {
    super(props);

    // If this is running in Teams, embed the specific Teams styling
    if (this._isRunningInTeams()) {
      import(/* webpackChunkName: 'LightTheme' */"./GraphCalendar.Teams.module.scss");

      if (this.props.teamsContext.app.theme === "dark") {
        import(/* webpackChunkName: 'DarkTheme' */"./GraphCalendar.Teams.Dark.module.scss");
      }
    }

    //Set language of the calendar (if language is unknown, use english as default)
    const allLanguages: Map<string, string> = this._createLangMap();
    if (this._isRunningInTeams()) {
      this.calendarLang = allLanguages.get(this.props.teamsContext.app.locale) || "en";
    } else {
      this.calendarLang = allLanguages.get(this.props.context.pageContext.cultureInfo.currentCultureName) || "en";
    }

    let resolvedGroupId = this.props.group && this.props.group.length ? this.props.group[0].id : "";
    if (resolvedGroupId) {
      const match = resolvedGroupId.match(/\w{8}-\w{4}-\w{4}-\w{4}-\w{12}/);
      resolvedGroupId = match ? match[0] : '';
    } else {
      resolvedGroupId = this._isRunningInTeams() ? this.props.teamsContext.team?.groupId : this.props.context.pageContext.site.group ? this.props.context.pageContext.site.group.id : "";
    }

    this.state = {
      events: [],
      height: this._calculateHeight(),
      currentActiveStartDate: undefined,
      currentActiveEndDate: undefined,
      isEventDetailsOpen: false,
      currentSelectedEvent: undefined,
      groupId: resolvedGroupId,
      tabType: this._isRunningInTeams() ? (this._isPersonalTab() ? TabType.PersonalTab : TabType.TeamsTab) : TabType.TeamsTab
    };
  }

  /**
   * When the component is initially mounted
   */
  public componentDidMount(): void {
    // Gets the calendar current Active dates
    const calendarStartDate = this.calendar.current?.getApi().view.activeStart;
    const calendarEndDate = this.calendar.current?.getApi().view.activeEnd;

    if (calendarStartDate && calendarEndDate) {
      // Loads the events
      this._loadEvents(calendarStartDate, calendarEndDate);
    }
  }
  
   /**
   * Renders the web part
   */
   public render(): React.ReactElement<IGraphCalendarProps> {
    return (
      <div className={styles.graphCalendar}>
        <FullCalendar
          ref={this.calendar}
          initialView="dayGridMonth"
          plugins={[dayGridPlugin]}
          windowResize={this._handleResize.bind(this)}
          datesSet={this._datesRender.bind(this)}
          eventClick={this._openEventPanel.bind(this)}
          height={this.state.height}
          events={this.state.events}
          locales={allLocales}
          locale={this.calendarLang} />
        {this.state.currentSelectedEvent &&
          <Panel
            isOpen={this.state.isEventDetailsOpen}
            type={PanelType.smallFixedFar}
            headerText={this.state.currentSelectedEvent ? this.state.currentSelectedEvent.title : ""}
            onDismiss={this._closeEventPanel.bind(this)}
            isLightDismiss={true}
            closeButtonAriaLabel={strings.Close}>
            <h3>{strings.StartTime}</h3>
            <span>{moment(this.state.currentSelectedEvent.start).format('MMMM Do YYYY [at] h:mm:ss a')}</span>
            <h3>{strings.EndTime}</h3>
            <span>{moment(this.state.currentSelectedEvent.end).format('MMMM Do YYYY [at] h:mm:ss a')}</span>
            {this.state.currentSelectedEvent.extendedProps && this.state.currentSelectedEvent.extendedProps.location &&
              <div>
                <h3>{strings.Location}</h3>
                <span>{this.state.currentSelectedEvent.extendedProps.location}</span>
              </div>
            }
            {this.state.currentSelectedEvent.extendedProps && this.state.currentSelectedEvent.extendedProps.body &&
              <div>
                <h3>{strings.Body}</h3>
                <span>{this.state.currentSelectedEvent.extendedProps.body}</span>
              </div>
            }
          </Panel>
        }
      </div>
    );
  }
  
  /**
   * Calculates the dynamic height of the surface to render.
   * Mainly used for Teams validation so it renders "full-screen" in Teams
   */
  private _calculateHeight(): number {
    if (this._isRunningInTeams()) {
      return window.innerHeight - 30;
    } else {
      return 600;
    }
  }

  /**
   * Validates if the current web part is running in Teams
   */
  private _isRunningInTeams(): microsoftTeams.app.Context {
    return this.props.teamsContext;
  }

  /**
   * Validates if the current web part is running in a Personal Tab
   */
  private _isPersonalTab(): boolean {
    let _isPersonalTab: boolean = false;

    if (this._isRunningInTeams() && !this.props.teamsContext.team?.internalId) {
      _isPersonalTab = true;
    }

    return _isPersonalTab;
  }

  /**
   * Handles the click event and opens the OUIF Panel
   * @param eventClickInfo The information about the selected event
   */
  private _openEventPanel(eventClickInfo: any): void {
    this.setState({
      isEventDetailsOpen: true,
      currentSelectedEvent: eventClickInfo.event
    });
  }

  /**
   * Handles the click event on the dismiss from the Panel and closes the OUIF Panel
   */
  private _closeEventPanel(): void {
    this.setState({
      isEventDetailsOpen: true,
      currentSelectedEvent: undefined
    });
  }

  /**
   * If the view changed, reload the events based on the active view
   * @param info Information about the current active view
   */
  private _datesRender(info: any): void {
    if (this.calendar.current) {

      // If the active view has changed  
      if ((this.state.currentActiveStartDate && this.state.currentActiveEndDate) && this.state.currentActiveStartDate.toString() !== info.view.activeStart.toString() && this.state.currentActiveEndDate.toString() !== info.view.activeEnd.toString()) {
        this._loadEvents(info.view.activeStart, info.view.activeEnd);
      }
    }
  }

  /**
   * Handles the resize event when in Microsoft Teams to ensure a proper responsive behaviour
   */
  private _handleResize(): void {
    if (this._isRunningInTeams()) {
      this.setState({
        height: window.innerHeight - 30
      });
    }
  }

  /**
   * Convert data to Array<EventInput>
   * @param data Events from API
   */
  private _transformEvents(data: any): Array<EventInput> {
    const events: Array<EventInput> = new Array<EventInput>();
    data.value.map((item: any) => {
      // Build a Timezone enabled Date
      const currentStartDate = moment.tz(item.start.dateTime, item.start.timeZone);
      const currentEndDate = moment.tz(item.end.dateTime, item.end.timeZone);

      // Adding all retrieved events to the result array
      events.push({
        id: item.id,
        title: item.subject,

        // If the event is an All Day event, add 1 day without Timezone to the start date
        start: !item.isAllDay ? currentStartDate.clone().tz(Intl.DateTimeFormat().resolvedOptions().timeZone).format() : moment(currentStartDate).add(1, 'd').toISOString(),

        // If the event is an All Day event, add 1 day without Timezone to the end date
        end: !item.isAllDay ? currentEndDate.clone().tz(Intl.DateTimeFormat().resolvedOptions().timeZone).format() : moment(currentEndDate).add(1, 'd').toISOString(),
        allDay: item.isAllDay,
        location: item.location.displayName,
        body: item.bodyPreview,
        type: item.type
      });
    });
    return events;
  }

  /**
   * Check if the recurring events need to be shown on the current state of the Calendar
   * If the range of the recurring event overlaps the range of the Calendar, then the event needs to be shown.
   * @param data All the recurrent (base) events ever made
   * @param startDate The first visible date on the calendar
   * @param endDate The last visible date on the calendar
   */
  private _filterRecEvents(data: any, startDate: Date, endDate: Date): Array<EventInput> {
    const events: Array<EventInput> = new Array<EventInput>();
    //Range of the Calendar
    const r1 = range(startDate, endDate);

    data.value.map((item: any) => {
      // Build a Timezone enabled Date
      const currentStartDate = moment.tz(item.start.dateTime, item.start.timeZone);
      const currentEndDate = moment.tz(item.end.dateTime, item.end.timeZone);

      const d1 = item.recurrence.range.startDate;
      const d2 = item.recurrence.range.endDate;
      const recStartDate = moment(d1).toDate();
      const recEndDate = moment(d2).toDate();

      //Range of the recurring event item
      const r2 = range(recStartDate, recEndDate);

      //Check if both ranges overlap
      if (!!r1.overlaps(r2)) {
        events.push({
          id: item.id,
          title: item.subject,
          // If the event is an All Day event, add 1 day without Timezone to the start date
          start: !item.isAllDay ? currentStartDate.clone().tz(Intl.DateTimeFormat().resolvedOptions().timeZone).format() : moment(currentStartDate).add(1, 'd').toISOString(),
          // If the event is an All Day event, add 1 day without Timezone to the end date
          end: !item.isAllDay ? currentEndDate.clone().tz(Intl.DateTimeFormat().resolvedOptions().timeZone).format() : moment(currentEndDate).add(1, 'd').toISOString(),
          allDay: item.isAllDay,
          location: item.location.displayName,
          body: item.bodyPreview,
          type: item.type
        });
      }

    });
    return events;
  }

  /**
   * Loads the Events based on the current state of the Calendar
   * @param startDate The first visible date on the calendar
   * @param endDate The last visible date on the calendar
   */
  private _loadEvents(startDate: Date, endDate: Date): void {

    // If a Group was found or running in the context of a Personal tab, execute the query. If not, do nothing.
    if (this.state.groupId || this.state.tabType === TabType.PersonalTab) {
      let events: Array<EventInput> = new Array<EventInput>();

      this.props.context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          let apiUrl: string = `/groups/${this.state.groupId}/events`;
          if (this._isPersonalTab()) {
            apiUrl = '/me/events';
          }

          client
            .api(apiUrl)
            .version("v1.0")
            .select('subject,start,end,location,bodyPreview,isAllDay,type')
            .filter(`start/dateTime ge '${startDate.toISOString()}' and end/dateTime le '${endDate.toISOString()}' and type eq 'singleInstance'`)
            .top(this.props.limit)
            .get((err, res) => {

              if (err) {
                console.error(err);
                return;
              }

              //Transform API data to Array<EventInput>
              events = this._transformEvents(res);

              if (this.props.showRecurrence) {
                //Get recurring events and merge with the other (standard) events
                this._getRecurringMaster(startDate, endDate).then((recData: Array<EventInput>) => {

                  if (recData.length > 0) {
                    this._getRecurrentEvents(recData, startDate, endDate).then((recEvents: Array<EventInput>) => {
                      this.setState({
                        events: [...recEvents, ...events].slice(0, this.props.limit),
                      });
                    });
                  } else {
                    this.setState({
                      events: events,
                    });
                  }

                });
              } else {
                //Show only (standard) events
                this.setState({
                  events: events,
                });
              }

              // Sets the state with current active calendar dates
              this.setState({
                currentActiveStartDate: startDate,
                currentActiveEndDate: endDate,
                currentSelectedEvent: undefined
              });
            });
        });
    }
  }

  /**
  * Get all recurrent events based on the current state of the Calendar
  * @param events All the recurrent base events
  * @param startDate The first visible date on the calendar
  * @param endDate The last visible date on the calendar
  */
  private _getRecurrentEvents(events: Array<EventInput>, startDate: Date, endDate: Date): Promise<Array<EventInput>> {
    return new Promise<Array<EventInput>>((resolve, reject) => {
      this.props.context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          let recEvents: Array<EventInput> = new Array<EventInput>();
          let count = 0;
          events.map((item: EventInput) => {
            let apiUrl: string = `/groups/${this.state.groupId}/events/${item.id}/instances?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`;
            if (this._isPersonalTab()) {
              apiUrl = `/me/events/${item.id}/instances?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`;
            }
            client
              .api(apiUrl)
              .version("v1.0")
              .select('subject,start,end,location,bodyPreview,isAllDay,type')
              .get((err, res) => {
                if (err) {
                  reject(err);
                }
                recEvents = recEvents.concat(this._transformEvents(res));
                count += 1;
                if (count === events.length) {
                  resolve(recEvents);
                }
              });
          });

        });
    });
  }

  /**
   * Get all recurrent (base) events ever created
   * Filter the base events based on the current state of the Calendar
   * @param startDate The first visible date on the calendar
   * @param endDate The last visible date on the calendar
   */
  private _getRecurringMaster(startDate: Date, endDate: Date): Promise<Array<EventInput>> {
    return new Promise<Array<EventInput>>((resolve, reject) => {
      this.props.context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          let apiUrl: string = `/groups/${this.state.groupId}/events`;
          if (this._isPersonalTab()) {
            apiUrl = '/me/events';
          }
          client
            .api(apiUrl)
            .version("v1.0")
            .select('subject,start,end,location,bodyPreview,isAllDay,type,recurrence')
            .filter(`type eq 'seriesMaster'`) //recurrening event is type 'seriesMaster'
            .get((err, res) => {
              if (err) {
                reject(err);
              }
              else {
                let recEvents: Array<EventInput> = new Array<EventInput>();
                recEvents = this._filterRecEvents(res, startDate, endDate);
                resolve(recEvents);
              }
            });
        });
    });
  }

   /**
   * Mapping for SharePoint languages with Fullcalendar languages
   */
  private _createLangMap(): Map<string, string> {
    const languages = new Map();

    languages.set("en-US", "en"); //English
    languages.set("ar-SA", "ar"); //Arabic
    languages.set("bg-BG", "bg"); //Bulgarian
    languages.set("ca-ES", "ca"); //Catalan
    languages.set("cs-CZ", "cs"); //Czech
    languages.set("da-DK", "da"); //Danish
    languages.set("de-DE", "de"); //German
    languages.set("el-GR", "el"); //Greek
    languages.set("es-ES", "es"); //Spanish
    languages.set("et-EE", "et"); //Estonian
    languages.set("fi-FI", "fi"); //Finish
    languages.set("fr-FR", "fr"); //French
    languages.set("he-IL", "he"); //Hebrew
    languages.set("hi-IN", "hi"); //Hindi
    languages.set("hr-HR", "hr"); //Croatian
    languages.set("hu-HU", "hu"); //Hungarian
    languages.set("it-IT", "it"); //Italian
    languages.set("kk-KZ", "kk"); //Kazakh
    languages.set("ko-KR", "ko"); //Korean
    languages.set("lt-LT", "lt"); //Lithuanian
    languages.set("lv-LV", "lv"); //Latvian
    languages.set("nb-NO", "nb"); //Norwegian
    languages.set("nl-NL", "nl"); //Dutch
    languages.set("pl-PL", "pl"); //Polish
    languages.set("pt-BR", "pt-br"); //Portugues (Brazil)
    languages.set("pt-PT", "pt"); //Portuguese (Portugal)
    languages.set("ro-RO", "ro"); //Romanian
    languages.set("ru-RU", "ru"); //Russian
    languages.set("sk-SK", "sk"); //Slovak
    languages.set("sl-SI", "sl"); //Slovenian
    languages.set("sr-Latn-CS", "sr-cyrl"); //Serbian
    languages.set("sv-SE", "sv"); //Swedish
    languages.set("th-TH", "th"); //Thai
    languages.set("tr-TR", "tr"); //Turkish
    languages.set("uk-UA", "uk"); //Ukrainian
    languages.set("zh-CN", "zh-cn"); //Chinese (Simplified)
    languages.set("zh-TW", "zh-tw"); //Chinese (Taiwan)

    return languages;
  }
}
