using System;
using Microsoft.SharePoint;
namespace BethesdaConsole
{
    class Program
    {
        static void Main()
        {
            try
            {
                //Console.Clear();
                //Console.WriteLine("Please input site url : ");
                string groupName;
                const string siteUrl = "http://tspsrvr";
                SPSecurity.RunWithElevatedPrivileges(delegate
                {
                    using (var site = new SPSite(siteUrl))
                    {
                        using (var web = site.OpenWeb())
                        {
                            var list = web.Lists["Course"];
                            foreach (SPListItem spListItem in list.Items)
                            {
                                try
                                {
                                    groupName = spListItem.Title;
                                    SPMember siteOwner = web.Site.RootWeb.SiteAdministrators[0];
                                    web.AllowUnsafeUpdates = true;
                                    web.SiteGroups.Add(groupName, siteOwner, web.Users[0], "");
                                    SPGroup wcmGroup = web.SiteGroups[groupName];
                                    SPRoleDefinition customRoleDefinition = web.RoleDefinitions["Read"];
                                    var roleAssignment = new SPRoleAssignment(wcmGroup);
                                    roleAssignment.RoleDefinitionBindings.Add(customRoleDefinition);
                                    web.RoleAssignments.Add(roleAssignment);
                                    wcmGroup.Update();
                                    web.Update();
                                    web.AllowUnsafeUpdates = false;

                                     // Email sending through programatically
                                    StringDictionary headers = new StringDictionary();

                                    headers.Add("from", "sender@domain.com");
                                    headers.Add("to", "praveen@tillidsoft.com");
                                    headers.Add("subject", "Welcome to the SharePoint");
                                    headers.Add("fAppendHtmlTag", "True"); //To enable HTML Tags

                                    System.Text.StringBuilder strMessage = new System.Text.StringBuilder();
                                    strMessage.Append("Message from CEO:");

                                    strMessage.Append("<span style='color:red;'> Make sure you have completed the survey! </span>");
                                    SPUtility.SendEmail(web, headers, strMessage.ToString());
                                }
                                catch (Exception)
                                {
                                    continue;
                                }
                            }
                        }
                    }
                });
            }
            catch (Exception ex) {
                    Console.WriteLine("Error:");
                    Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Process Completed");
                Console.ReadKey();
            }
        }

        internal SPListItemCollection GetListItemCollection(SPList spList, string key, string value)
        {
            // Return list item collection based on the lookup field

            SPField spField = spList.Fields[key];
            var query = new SPQuery
            {
                Query = @"<Where>
                        <Eq>
                            <FieldRef Name='" + spField.InternalName + @"'/><Value Type='" + spField.Type.ToString() + @"'>" + value + @"</Value>
                        </Eq>
                        </Where>"
            };

            return spList.GetItems(query);
        }

        internal SPListItemCollection GetListItemCollection(SPList spList, string keyOne, string valueOne, string keyTwo, string valueTwo)
        {
            // Return list item collection based on the lookup field

            SPField spFieldOne = spList.Fields[keyOne];
            SPField spFieldTwo = spList.Fields[keyTwo];
            var query = new SPQuery
            {
                Query = @"<Where>
                          <And>
                                <Eq>
                                    <FieldRef Name=" + spFieldOne.InternalName + @" />
                                    <Value Type=" + spFieldOne.Type.ToString() + ">" + valueOne + @"</Value>
                                </Eq>
                                <Eq>
                                    <FieldRef Name=" + spFieldTwo.InternalName + @" />
                                    <Value Type=" + spFieldTwo.Type.ToString() + ">" + valueTwo + @"</Value>
                                </Eq>
                          </And>
                        </Where>"
            };

            return spList.GetItems(query);
        }

        internal SPListItemCollection GetListItemCollection(SPList spList, string keyOne, string valueOne, string keyTwo, string valueTwo, string keyThree, string valueThree)
        {
            // Return list item collection based on the lookup field

            SPField spFieldOne = spList.Fields[keyOne];
            SPField spFieldTwo = spList.Fields[keyTwo];
            SPField spFieldThree = spList.Fields[keyThree];
            var query = new SPQuery
            {
                Query = @"<Where>
                          <And>
                             <And>
                                <Eq>
                                   <FieldRef Name=" + spFieldOne.InternalName + @" />
                                   <Value Type=" + spFieldOne.Type.ToString() + @">" + valueOne + @"</Value>
                                </Eq>
                                <Eq>
                                   <FieldRef Name=" + spFieldTwo.InternalName + @" />
                                   <Value Type=" + spFieldTwo.Type.ToString() + @">" + valueTwo + @"</Value>
                                </Eq>
                             </And>
                             <Eq>
                                <FieldRef Name=" + spFieldThree.InternalName + @" />
                                <Value Type=" + spFieldThree.Type.ToString() + @">" + valueThree + @"</Value>
                             </Eq>
                          </And>
                       </Where>"
            };

            return spList.GetItems(query);
        }

        internal SPListItemCollection GetListItemCollection(SPList spList, string keyOne, string valueOne, string keyTwo, string valueTwo, string keyThree, string valueThree, string keyFour, string valueFour)
        {
            // Return list item collection based on the lookup field

            SPField spFieldOne = spList.Fields[keyOne];
            SPField spFieldTwo = spList.Fields[keyTwo];
            SPField spFieldThree = spList.Fields[keyThree];
            SPField spFieldFour = spList.Fields[keyFour];
            var query = new SPQuery
            {
                Query = @"<Where>
                          <And>
                             <And>
                                <And>
                                <Eq>
                                   <FieldRef Name=" + spFieldOne.InternalName + @" />
                                   <Value Type=" + spFieldOne.Type.ToString() + @">" + valueOne + @"</Value>
                                </Eq>
                                <Eq>
                                   <FieldRef Name=" + spFieldTwo.InternalName + @" />
                                   <Value Type=" + spFieldTwo.Type.ToString() + @">" + valueTwo + @"</Value>
                                </Eq>
                             </And>
                             <Eq>
                                <FieldRef Name=" + spFieldThree.InternalName + @" />
                                <Value Type=" + spFieldThree.Type.ToString() + @">" + valueThree + @"</Value>
                             </Eq>
                          </And>
                             <Eq>
                                <FieldRef Name=" + spFieldFour.InternalName + @" />
                                <Value Type=" + spFieldFour.Type.ToString() + @">" + valueFour + @"</Value>
                             </Eq>
                          </And>
                       </Where>"
            };

            return spList.GetItems(query);
        }
    }
}
