using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ResourceCalendarNet
{
    public class ResourceCalendarClass
    {
        public EntityCollection getResourceCalendarEntity(IOrganizationService service, Guid ResourceID, DateTime StartDate)
        {
            var QEbookableresource = new QueryExpression("bookableresource") { NoLock = true };
            QEbookableresource.ColumnSet.AddColumns("calendarid");
            FilterExpression ResourceFilter = new FilterExpression(LogicalOperator.Or);
            ResourceFilter.AddCondition("bookableresourceid", ConditionOperator.Equal, ResourceID);

            QEbookableresource.Criteria.AddFilter(ResourceFilter);
            var QEbookableresource_calendar = QEbookableresource.AddLink("calendar", "calendarid", "calendarid");
            var QEbookableresource_calendar_calendarrule = QEbookableresource_calendar.AddLink("calendarrule", "calendarid", "calendarid");
            QEbookableresource_calendar_calendarrule.EntityAlias = "calendardate";
            QEbookableresource_calendar_calendarrule.Columns.AddColumns("name", "starttime", "pattern", "calendarruleid");

            QEbookableresource_calendar_calendarrule.LinkCriteria.AddCondition(new ConditionExpression("starttime", ConditionOperator.On, StartDate));

            var QEbookableresource_calendar_calendarrule_calendar = QEbookableresource_calendar_calendarrule.AddLink("calendar", "innercalendarid", "calendarid");
            QEbookableresource_calendar_calendarrule_calendar.EntityAlias = "innercalendar";
            QEbookableresource_calendar_calendarrule_calendar.Columns.AddColumns("calendarid", "name");
            var QEbookableresource_calendar_calendarrule_calendar_calendarrule = QEbookableresource_calendar_calendarrule_calendar.AddLink("calendarrule", "calendarid", "calendarid");
            QEbookableresource_calendar_calendarrule_calendar_calendarrule.EntityAlias = "calendartime";
            QEbookableresource_calendar_calendarrule_calendar_calendarrule.Columns.AddColumns("description", "starttime", "duration", "offset", "description", "calendarid", "name", "timecode", "subcode");
            return service.RetrieveMultiple(QEbookableresource);
        }

        public void CreateCalendar(IOrganizationService service, DateTime CalendarDateTime, Guid bookableresourceid, int timecode, int subcode, int duration, int offset)
        {
            var bookableresouceresp = service.Retrieve("bookableresource", bookableresourceid, new ColumnSet("calendarid"));
            EntityReference CalendarID = (EntityReference)bookableresouceresp["calendarid"];
            var userCalendarEntity = service.Retrieve(CalendarID.LogicalName, CalendarID.Id, new ColumnSet(true));
            var calendarRules = (EntityCollection)userCalendarEntity["calendarrules"];



            Entity newInnerCalendar = new Entity("calendar");
            userCalendarEntity.Attributes["type"] = new OptionSetValue(-1);
            newInnerCalendar.Attributes["businessunitid"] = new EntityReference("businessunit", ((Microsoft.Xrm.Sdk.EntityReference)(userCalendarEntity["businessunitid"])).Id);
            Guid innerCalendarId = service.Create(newInnerCalendar);




            newInnerCalendar = service.Retrieve(newInnerCalendar.LogicalName, innerCalendarId, new ColumnSet(true));
            EntityCollection innerCalendarRules = (EntityCollection)newInnerCalendar["calendarrules"];
            Entity calendarRule = new Entity("calendarrule");
            calendarRule.Attributes["duration"] = duration;
            calendarRule.Attributes["extentcode"] = 1;
            calendarRule.Attributes["pattern"] = "FREQ=DAILY;INTERVAL=1;COUNT=1";
            calendarRule.Attributes["rank"] = 0;
            calendarRule.Attributes["timezonecode"] = 105;
            calendarRule.Attributes["innercalendarid"] = new EntityReference("calendar", innerCalendarId);
            calendarRule.Attributes["starttime"] = CalendarDateTime;
            calendarRules.Entities.Add(calendarRule);
            userCalendarEntity.Attributes["calendarrules"] = calendarRules;
            service.Update(userCalendarEntity);



            Entity workingHourcalendarRule = new Entity("calendarrule");
            workingHourcalendarRule.Attributes["duration"] = duration;
            workingHourcalendarRule.Attributes["effort"] = 1.0;
            workingHourcalendarRule.Attributes["issimple"] = true;
            calendarRule.Attributes["starttime"] = CalendarDateTime;



            workingHourcalendarRule.Attributes["offset"] = offset; //to indicate start time is 8 AM
            workingHourcalendarRule.Attributes["rank"] = 0;



            //Type of calendar rule such as working hours, break, holiday, or time off. 0 for working hours
            workingHourcalendarRule.Attributes["timecode"] = timecode;
            workingHourcalendarRule.Attributes["subcode"] = subcode;
            workingHourcalendarRule.Attributes["timezonecode"] = 105;
            workingHourcalendarRule.Attributes["calendarid"] = new EntityReference("calendar", innerCalendarId);
            innerCalendarRules.Entities.Add(workingHourcalendarRule);
            service.Update(new Entity(newInnerCalendar.LogicalName, newInnerCalendar.Id)
            {
                Attributes = new AttributeCollection()
                {
                    new KeyValuePair<string, object>("calendarrules", innerCalendarRules)
                }
            });
        }
    }
}
