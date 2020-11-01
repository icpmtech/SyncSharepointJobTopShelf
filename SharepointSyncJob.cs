using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using log4net;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using Quartz;
using SyncUserProfilesToListContatos;
using List = Microsoft.SharePoint.Client.List;
using ListItem = Microsoft.SharePoint.Client.ListItem;
using User = Microsoft.SharePoint.Client.User;

namespace SyncUserProfilesToListContatos
{
    /// <summary>
    /// Sample Job that executes based on schedule defined in quartzjobconfig.xml
    /// </summary>
    public class SharepointSyncJob : IJob
    {
        
        
        private readonly ILog _log;
        private readonly Configuration _configuration;

        public SharepointSyncJob(ILog log, IConfigurationProvider configurationProvider)
        {
            _log = log;
            _configuration = configurationProvider.GetConfiguration();


        }
        private void Init()
        {
           //var Users=  GetAllUsers();
           // Users.Wait();
            _log.Info("Início do processo: SharepointSyncJob");
            var listItemListaTelefonicaProperties = new List<ItemListaTelefonicaProperties>();
            using (ClientContext tenantContext = new ClientContext(_configuration.TenantUrl))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _configuration.Password.ToCharArray())
                    passWord.AppendChar(c);
                tenantContext.Credentials = new SharePointOnlineCredentials(_configuration.EmailAccount, passWord);
                PeopleManager peopleManager = new PeopleManager(tenantContext);

                UserCollection users = tenantContext.Web.SiteUsers;
                tenantContext.Load(users);
                tenantContext.ExecuteQuery();
                StringBuilder items = new StringBuilder();
                string[] userProfileProperties =   {
                    UserProfileProperties.FirstName,
                    UserProfileProperties.LastName,
                    UserProfileProperties.PreferredName,
                    UserProfileProperties.WorkEmail,
                    UserProfileProperties.Pager,
                    UserProfileProperties.Department,
                    UserProfileProperties.CellPhone,
                    UserProfileProperties.PictureURL,
                    UserProfileProperties.Office,
                    UserProfileProperties.Fax,
                    UserProfileProperties.WorkPhone,
                     UserProfileProperties.IpPhone,
                };

                foreach (string propertyKey in userProfileProperties)
                {
                    items.Append(propertyKey);
                    items.Append(",");
                }
                items.AppendLine();
                _log.Info("A ler os Users profile: SharepointSyncJob");
                foreach (User user in users)
                {
                    try
                    {
                        if (user.PrincipalType != Microsoft.SharePoint.Client.Utilities.PrincipalType.User) continue;

                        UserProfilePropertiesForUser userProfilePropertiesForUser = new UserProfilePropertiesForUser(tenantContext, user.LoginName, userProfileProperties);
                        var profileProperties = peopleManager.GetUserProfilePropertiesFor(userProfilePropertiesForUser);
                        tenantContext.Load(userProfilePropertiesForUser);
                        tenantContext.ExecuteQuery();
                        var listProfileProperties = profileProperties.ToList();
                        if (listProfileProperties.Any())
                        {
                            ItemListaTelefonicaProperties itemListaTelefonicaProperties = new ItemListaTelefonicaProperties();

                            itemListaTelefonicaProperties.FirstName = listProfileProperties[0];
                            itemListaTelefonicaProperties.LastName = listProfileProperties[1];
                            itemListaTelefonicaProperties.PreferredName = listProfileProperties[2];
                            itemListaTelefonicaProperties.Email = user.Email;
                            itemListaTelefonicaProperties.Pager = listProfileProperties[4];
                            itemListaTelefonicaProperties.Department = listProfileProperties[5];
                            itemListaTelefonicaProperties.CellPhone = listProfileProperties[6];
                            itemListaTelefonicaProperties.PictureURL = listProfileProperties[7];
                            itemListaTelefonicaProperties.Office = listProfileProperties[8];
                            itemListaTelefonicaProperties.WorkFax = listProfileProperties[9];
                            itemListaTelefonicaProperties.WorkPhone = listProfileProperties[10];
                            itemListaTelefonicaProperties.IpPhone = listProfileProperties[11];
                            listItemListaTelefonicaProperties.Add(itemListaTelefonicaProperties);
                        }
                    }
                    catch (Exception ex)
                    {
                        _log.Error($"-> users no processo: SharepointSyncJob: {ex}", ex);
                    }

                }

                InsertOrUpdateContactoToListContactos(listItemListaTelefonicaProperties);
            }
            _log.Info("Fim do processo: SharepointSyncJob");
        }

        

        private void InsertOrUpdateContactoToListContactos(List<ItemListaTelefonicaProperties> itemsItemListaTelefonicaProperties)
        {
            _log.Info("Inicio do metodo -> InsertOrUpdateContactoToListContactos no processo: SharepointSyncJob");
            using (ClientContext siteContexto = new ClientContext(_configuration.TenantUrl))
            {
                SecureString passWord = new SecureString();
                foreach (char c in _configuration.Password.ToCharArray())
                    passWord.AppendChar(c);
                siteContexto.Credentials = new SharePointOnlineCredentials(_configuration.EmailAccount, passWord);
                if (siteContexto.Web.ListExists(_configuration.ListTitle))
                {
                    Microsoft.SharePoint.Client.List listaTelefonica = siteContexto.Web.Lists.GetByTitle(_configuration.ListTitle);
                    try
                    {
                        foreach (var itemListaTelefonicaProperties in itemsItemListaTelefonicaProperties)
                        {
                            if (itemListaTelefonicaProperties.Pager == "1")
                            {
                                var id = GetIdByTitle(itemListaTelefonicaProperties.PreferredName);
                                if (id == null)
                                {
                                    _log.Info($"Criar novo contato -> CreateNovoContacto no processo: SharepointSyncJob, Valor: {itemListaTelefonicaProperties}");
                                    CreateNovoContacto(siteContexto, listaTelefonica, itemListaTelefonicaProperties);
                                }
                                else
                                {
                                    _log.Info($"Update Contacto -> UpdateContacto no processo: SharepointSyncJob, Valor: {itemListaTelefonicaProperties}");
                                    UpdateContacto(siteContexto, listaTelefonica, itemListaTelefonicaProperties, id.Value);
                                }

                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        _log.Error($"-> InsertOrUpdateContactoToListContactos no processo: SharepointSyncJob: {ex}", ex);
                    }
                }
            }
            _log.Info("Fim do metodo -> InsertOrUpdateContactoToListContactos no processo: SharepointSyncJob");

        }

        private void UpdateContacto(ClientContext siteContexto, Microsoft.SharePoint.Client.List listaTelefonica, ItemListaTelefonicaProperties itemListaTelefonicaProperties, int id)
        {
            try
            {


                Microsoft.SharePoint.Client.ListItem updateItem = listaTelefonica.GetItemById(id);
                updateItem["Title"] = itemListaTelefonicaProperties.PreferredName;
                updateItem["FirstName"] = itemListaTelefonicaProperties.FirstName;
                updateItem["LastName"] = itemListaTelefonicaProperties.LastName;
                updateItem["Email"] = itemListaTelefonicaProperties.Email;
                updateItem["Department"] = itemListaTelefonicaProperties.Department;
                updateItem["CellPhone"] = itemListaTelefonicaProperties.CellPhone;
                updateItem["PictureURL"] = itemListaTelefonicaProperties.PictureURL;
                updateItem["Office"] = itemListaTelefonicaProperties.Office;
                updateItem["WorkFax"] = itemListaTelefonicaProperties.WorkFax;
                updateItem["WorkPhone"] = itemListaTelefonicaProperties.WorkPhone;
                updateItem["WorkState"] = itemListaTelefonicaProperties.WorkState;
                updateItem["WorkCity"] = itemListaTelefonicaProperties.WorkCity;
                updateItem["WorkCountry"] = itemListaTelefonicaProperties.WorkCountry;
                updateItem["WorkZip"] = itemListaTelefonicaProperties.WorkZip;
                updateItem["IpPhone"] = itemListaTelefonicaProperties.IpPhone;
                updateItem["WorkAddress"] = itemListaTelefonicaProperties.WorkAddress;
                updateItem.Update();
                siteContexto.ExecuteQuery();
            }
            catch (Exception ex)
            {

                _log.Error($"-> UpdateContacto no processo: SharepointSyncJob: {ex}", ex);
            }
        }

        private int? GetIdByTitle(string title)
        {
            try
            {
                using (ClientContext clientContext = new ClientContext(_configuration.TenantUrl))
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in _configuration.Password.ToCharArray())
                        passWord.AppendChar(c);
                    clientContext.Credentials = new SharePointOnlineCredentials(_configuration.EmailAccount, passWord);
                    List list = clientContext.Web.Lists.GetByTitle(_configuration.ListTitle);
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = string.Format($@"
                        <View>	                
                        <Query>
                        <Where>	     
                        <Eq>	          
                        <FieldRef Name='Title' />
                        <Value Type='Text'>{title}</Value>	 
                        </Eq>
                        </Where>
                        </Query>
                        <ViewFields>
                        <FieldRef Name='ID'/>
                        <FieldRef Name='Title'/>
                        </ViewFields>
 <RowLimit>1</RowLimit>
</View>");
                    // reads the item with the specified caml query
                    ListItemCollection itemCollection = list.GetItems(camlQuery);
                    // tells the ClientContext object to load the collect, 	// while we define the fields that should be returned like:	// Id, Title, SomeDate, OtherField and Modified 
                    clientContext.Load(itemCollection,
                        items => items.Include(
                        item => item.Id,
                        item => item["Title"]));
                    // executes everything that was loaded before (the itemCollection)
                    clientContext.ExecuteQuery();
                    // iterates the list printing the title and another field of each list item
                    foreach (ListItem item in itemCollection)
                    {
                        return item.Id;
                    }
                }
            }
            catch (Exception ex)
            {

                _log.Error($"-> GetIdByTitle no processo: SharepointSyncJob: {ex}", ex);
            }

            return null;
        }

        private void CreateNovoContacto(ClientContext siteContexto, List listaTelefonica, ItemListaTelefonicaProperties itemListaTelefonicaProperties)
        {
            try
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem addItem = listaTelefonica.AddItem(itemCreateInfo);
                addItem["Title"] = itemListaTelefonicaProperties.PreferredName;
                addItem["FirstName"] = itemListaTelefonicaProperties.FirstName;
                addItem["LastName"] = itemListaTelefonicaProperties.LastName;
                addItem["Email"] = itemListaTelefonicaProperties.Email;
                addItem["PagerNumber"] = itemListaTelefonicaProperties.Pager;
                addItem["Department"] = itemListaTelefonicaProperties.Department;
                addItem["CellPhone"] = itemListaTelefonicaProperties.CellPhone;
                addItem["PictureURL"] = itemListaTelefonicaProperties.PictureURL;
                addItem["Office"] = itemListaTelefonicaProperties.Office;
                addItem["WorkFax"] = itemListaTelefonicaProperties.WorkFax;
                addItem["WorkFax"] = itemListaTelefonicaProperties.WorkFax;
                addItem["WorkPhone"] = itemListaTelefonicaProperties.WorkPhone;
                addItem["WorkState"] = itemListaTelefonicaProperties.WorkState;
                addItem["WorkCity"] = itemListaTelefonicaProperties.WorkCity;
                addItem["WorkCountry"] = itemListaTelefonicaProperties.WorkCountry;
                addItem["WorkZip"] = itemListaTelefonicaProperties.WorkZip;
                addItem["IpPhone"] = itemListaTelefonicaProperties.IpPhone;
                addItem["WorkAddress"] = itemListaTelefonicaProperties.WorkAddress;
                addItem.Update();
                siteContexto.ExecuteQuery();
            }
            catch (Exception ex)
            {
                _log.Error($"-> CreateNovoContacto no processo: SharepointSyncJob: {ex}", ex);

            }


        }

        async Task IJob.Execute(IJobExecutionContext context)
        {
            _log.Info("Execute Job is working");

            // Write background job logic. All the business logic goes here.

            await Task.Run(() => Init());
        }
    }
}
