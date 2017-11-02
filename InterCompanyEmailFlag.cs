using System;
using System.Text;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;


namespace InterCompanyEmailPlugin
{
    public class InterCompanyEmailFlag : IPlugin
    {
        /// <remarks>
        /// Strava Technology Group
        /// www.stravatechgroup.com
        /// A plug-in that checks to see if all email addresses of an email are all of the company domain
		/// Flags the email as Inter-Company
        /// This field could then be used to bulk delete from the system
        /// fields used:
        /// Custom Entity: new_companydomain; field: new_companyemaildomain
        /// Email Entity: new_intercompany
        /// Need to create Plugin steps for both create and update of email
        /// </remarks>
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
            context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];

                // Verify that the target entity represents an email.
                // If not, this plug-in was not registered correctly.
                if (entity.LogicalName != "email")
                    return;

                try
                {
                    //variables
                    Int32 iTotalActivityPartyCount = 0;
                    Int32 iInterCompanyActivityPartyCountContact = 0;
                    Int32 iInterCompanyActivityPartyCountUser = 0;
                    Int32 iInterCompanyActivityPartyCountUnresolved = 0;
                    Int32 iInterCompanyActivityPartyCountTotal = 0;

                    // Obtain the organization service reference.
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    Entity myEntity = new Entity(context.PrimaryEntityName);
                    myEntity = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet("statecode"));
                    
                        Guid EMailId = context.PrimaryEntityId;

                        //Find Domain.com from Company Email Domain record
                        string fetchDomain = @"
                                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='new_companyemaildomain'>
                                    <attribute name='new_companyemaildomain' />
                                    <filter type='and'>
                                      <condition attribute='new_companyemaildomain' operator='not-null' />
                                    </filter>
                                  </entity>
                                </fetch>";



                        EntityCollection EmailDomain = service.RetrieveMultiple(new FetchExpression(fetchDomain));
                        if (EmailDomain.Entities.Count == 1)
                        {
                            foreach (var childEntityD in EmailDomain.Entities)
                            {
                                string Domain = childEntityD.Attributes.Contains("new_companyemaildomain") ? ((string)childEntityD.Attributes["new_companyemaildomain"]) : "";

                                //Find Total Count of Activity Parties for current email
                                string fetchTotal = @"
                                <fetch version='1.0' output-format='xml-platform' mapping='logical' aggregate='true'>
                                  <entity name='activityparty'>
                                    <attribute name='activitypartyid' alias='activitypartycount' aggregate='count' />
                                    <filter type='and'>
                                      <condition attribute='activityid' operator='eq' value='{" + EMailId + @"}' />
                                      <condition attribute='participationtypemask' operator='in' >
                                        <value>1</value>
                                        <value>2</value>
                                        <value>3</value>
                                        <value>4</value>
                                      </condition>
                                    </filter>
                                  </entity>
                                </fetch>";

                                EntityCollection TotalActivityPartyCount = service.RetrieveMultiple(new FetchExpression(fetchTotal));
                                if (TotalActivityPartyCount.Entities.Count == 1)
                                {
                                    foreach (var childEntity in TotalActivityPartyCount.Entities)
                                    {
                                        iTotalActivityPartyCount = childEntity.Attributes.Contains("activitypartycount") ? ((int)((AliasedValue)childEntity.Attributes["activitypartycount"]).Value) : 0;
                                    }
                                }

                                //Find Total Count of Activity Parties (contacts) where company domain is used
                                string fetchCompanyContact = @"
                                <fetch version='1.0' output-format='xml-platform' mapping='logical' aggregate='true'>
                                  <entity name='activityparty'>
                                    <attribute name='activitypartyid' alias='activitypartycount' aggregate='count' />
                                    <filter type='and'>
                                      <condition attribute='activityid' operator='eq' value='{" + EMailId + @"}' />
                                      <condition attribute='partyobjecttypecode' operator='eq' value='2' />
                                      <condition attribute='participationtypemask' operator='in' >
                                        <value>1</value>
                                        <value>2</value>
                                        <value>3</value>
                                        <value>4</value>
                                      </condition>
                                    </filter>
                                  <link-entity name='contact' from='contactid' to='partyid' alias='ab' link-type='inner'>
                                    <filter type='and'>
                                      <condition attribute='emailaddress1' operator='like' value='%" + Domain + @"' />
                                    </filter>
                                  </link-entity>
                                  </entity>
                                </fetch>";



                                EntityCollection InterCompanyActivityPartyCountContact = service.RetrieveMultiple(new FetchExpression(fetchCompanyContact));
                                if (InterCompanyActivityPartyCountContact.Entities.Count == 1)
                                {
                                    foreach (var childEntityC in InterCompanyActivityPartyCountContact.Entities)
                                    {
                                        iInterCompanyActivityPartyCountContact = childEntityC.Attributes.Contains("activitypartycount") ? ((int)((AliasedValue)childEntityC.Attributes["activitypartycount"]).Value) : 0;
                                    }
                                }

                                //Find Total Count of Activity Parties (users) where company domain is used
                                string fetchCompanyUser = @"
                                <fetch version='1.0' output-format='xml-platform' mapping='logical' aggregate='true'>
                                  <entity name='activityparty'>
                                    <attribute name='activitypartyid' alias='activitypartycount' aggregate='count' />
                                    <filter type='and'>
                                      <condition attribute='activityid' operator='eq' value='{" + EMailId + @"}' />
                                      <condition attribute='partyobjecttypecode' operator='eq' value='8' />
                                      <condition attribute='participationtypemask' operator='in' >
                                        <value>1</value>
                                        <value>2</value>
                                        <value>3</value>
                                        <value>4</value>
                                      </condition>
                                    </filter>
                                  <link-entity name='systemuser' from='systemuserid' to='partyid' alias='ab' link-type='inner'>
                                    <filter type='and'>
                                      <condition attribute='internalemailaddress' operator='like' value='%" + Domain + @"' />
                                    </filter>
                                  </link-entity>
                                  </entity>
                                </fetch>";

                                EntityCollection InterCompanyActivityPartyCountUser = service.RetrieveMultiple(new FetchExpression(fetchCompanyUser));
                                if (InterCompanyActivityPartyCountUser.Entities.Count == 1)
                                {
                                    foreach (var childEntityU in InterCompanyActivityPartyCountUser.Entities)
                                    {
                                        iInterCompanyActivityPartyCountUser = childEntityU.Attributes.Contains("activitypartycount") ? ((int)((AliasedValue)childEntityU.Attributes["activitypartycount"]).Value) : 0;
                                    }
                                }

                            //Find Total Count of Activity Parties (unresolved) where company domain is used
                            string fetchCompanyUnresolved = @"
                                <fetch version='1.0' output-format='xml-platform' mapping='logical' aggregate='true'>
                                  <entity name='activityparty'>
                                    <attribute name='activitypartyid' alias='activitypartycount' aggregate='count' />
                                    <filter type='and'>
                                      <condition attribute='activityid' operator='eq' value='{" + EMailId + @"}' />
                                      <condition attribute='partyid' operator='null' />
                                      <condition attribute='addressused' operator='like' value='%" + Domain + @"'/>
                                      <condition attribute='participationtypemask' operator='in' >
                                        <value>1</value>
                                        <value>2</value>
                                        <value>3</value>
                                        <value>4</value>
                                      </condition>
                                    </filter>
                                  </entity>
                                </fetch>";

                            EntityCollection InterCompanyActivityPartyCountUnresolved = service.RetrieveMultiple(new FetchExpression(fetchCompanyUnresolved));
                            if (InterCompanyActivityPartyCountUnresolved.Entities.Count == 1)
                            {
                                foreach (var childEntityUR in InterCompanyActivityPartyCountUnresolved.Entities)
                                {
                                    iInterCompanyActivityPartyCountUnresolved = childEntityUR.Attributes.Contains("activitypartycount") ? ((int)((AliasedValue)childEntityUR.Attributes["activitypartycount"]).Value) : 0;
                                }
                            }

                            iInterCompanyActivityPartyCountTotal = iInterCompanyActivityPartyCountContact + iInterCompanyActivityPartyCountUser + iInterCompanyActivityPartyCountUnresolved;

                                if (iTotalActivityPartyCount == iInterCompanyActivityPartyCountTotal)
                                {
                                    //Update email tag
                                    Entity UpdatedEntity = new Entity("email");
                                    UpdatedEntity.Id = context.PrimaryEntityId;
                                    UpdatedEntity.Attributes["new_intercompany"] = true;
                                    service.Update(UpdatedEntity);
                                }
                            }
                        }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the FollowupPlugin plug-in.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("FollowupPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
