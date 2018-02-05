using System;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Core.WebApi;
using ClosedXML.Excel;
using System.Data;
using System.Collections.Generic;
using Microsoft.VisualStudio.Services.Common;
using System.IO;
using System.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ProjectAdminPoC.Classes;

namespace ProjectAdminPoC
{
    class Program
    {
        static void Main(string[] args)
        {
            //Get Office VSTS Projects
            IEnumerable<TeamProjectReference> projects = GetVstsProjects();

            //Access VSTS Groups
            TfsConfigurationServer tcs = new TfsConfigurationServer(new Uri(ConfigurationManager.AppSettings["Uri"]));
            IIdentityManagementService ims = tcs.GetService<IIdentityManagementService>();

            //Create Table for Storing Excel Data
            DataTable admins = new DataTable();
            admins.Clear();
            admins.Columns.Add("ProjectName");
            admins.Columns.Add("Email");

            //List for storing json strings
            List<Project> projectList = new List<Project>();

            //Iterate through projects to get admins
            foreach (TeamProjectReference project in projects)
            {
                string projectInfo = $"[{project.Name}]\\{ConfigurationManager.AppSettings["Group"]}";
                TeamFoundationIdentity tfi = ims.ReadIdentity(IdentitySearchFactor.AccountName, projectInfo, MembershipQuery.Direct, ReadIdentityOptions.None);
                List<TeamFoundationIdentity> ids = new List<TeamFoundationIdentity>();

                foreach (IdentityDescriptor identity in tfi.Members)
                {
                    try
                    {
                        TeamFoundationIdentity group = ims.ReadIdentity(identity,
                            MembershipQuery.ExpandedDown, ReadIdentityOptions.None);
                        GetAllProjectAdmins(group, ims, ref ids);
                    }
                    catch
                    {
                        TeamFoundationIdentity single = ims.ReadIdentity(identity, MembershipQuery.None, ReadIdentityOptions.None);
                        ids.Add(single);
                    }
                }

                Console.WriteLine($"Members of {projectInfo}");
                Project proj = new Project { Name = project.Name };
                proj.Administrators = new List<string>();

                foreach (TeamFoundationIdentity identity in ids)
                {
                    if (!proj.Administrators.Contains(identity.UniqueName))
                    {
                        //Add admin to DataTable
                        DataRow admin = admins.NewRow();
                        admin["ProjectName"] = project.Name;
                        admin["Email"] = identity.UniqueName;
                        admins.Rows.Add(admin);

                        //Add admin to Project Admin List
                        proj.Administrators.Add(identity.UniqueName);

                        Console.WriteLine(identity.UniqueName);
                    }
                }

                projectList.Add(proj);
            }

            //SaveAsExcel(admins);
            SaveAsJson(projectList);
        }

        public static IEnumerable<TeamProjectReference> GetVstsProjects()
        {
            VssBasicCredential credentials = new VssBasicCredential("", ConfigurationManager.AppSettings["PersonalAccessToken"]);

            IEnumerable<TeamProjectReference> projects;
            using (ProjectHttpClient projectHttpClient = new ProjectHttpClient(new Uri(ConfigurationManager.AppSettings["Uri"]), credentials))
            {
                projects = projectHttpClient.GetProjects().Result;
            }
            return projects;
        }

        public static void GetAllProjectAdmins(TeamFoundationIdentity identity, IIdentityManagementService ims, ref List<TeamFoundationIdentity> ids)
        {
            if (identity.IsContainer)
            {
                TeamFoundationIdentity[] groupMembers;

                try
                {
                    groupMembers = ims.ReadIdentities(identity.Members, MembershipQuery.Expanded,
                        ReadIdentityOptions.None);
                    foreach (TeamFoundationIdentity tfi in groupMembers)
                    {
                        GetAllProjectAdmins(tfi, ims, ref ids);
                    }
                }
                catch (Exception e)
                {
                    groupMembers = ims.ReadIdentities(identity.Members, MembershipQuery.None, ReadIdentityOptions.None);
                    ids.AddRange(groupMembers);
                }

            }
            else
            {
                ids.Add(identity);
            }
        }

        public static MemoryStream SaveWorkbookToMemoryStream(XLWorkbook workbook)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = false, GenerateCalculationChain = false, ValidatePackage = false });
                return stream;
            }
        }

        public static void SaveAsExcel(DataTable dt)
        {
            //Save to Excel
            FileInfo fileInfo = new FileInfo(ConfigurationManager.AppSettings["ExcelFilePath"]);
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, ConfigurationManager.AppSettings["WorksheetName"]);
            using (MemoryStream memoryStream = SaveWorkbookToMemoryStream(wb))
            {
                File.WriteAllBytes(fileInfo.FullName, memoryStream.ToArray());
            }
            wb.Dispose();
        }

        public static void SaveAsJson(List<Project> projectList)
        {
            string json = JsonConvert.SerializeObject(projectList, Formatting.Indented);
            File.WriteAllText($@"{ConfigurationManager.AppSettings["JsonFilePath"]}", json);
        }
    }
}
