using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Users_Permission_Comparer.Connection_Adapter;

namespace Users_Permission_Comparer.Comparer
{
    public class PermissionComparer
    {
        private Logger.Logger _Logger;
        private ConnectionAdapter _connectionAdapter;

        public PermissionComparer(ConnectionAdapter connectionAdapter, Logger.Logger Logger)
        {
            _connectionAdapter = connectionAdapter;
            _Logger = Logger;
        }

        /// <summary>
        /// Write into a file the users own permissions, then which they both have
        /// and lastly the permissions which only one has
        /// </summary>
        /// <param name="firstUser">fist selected user, which comes from the config</param>
        /// <param name="secondUser">second selected user, which comes from the config</param>
        public void WriteToFileUsersPermissions(string firstUser, string secondUser)
        {
            ICollection<string> firstUserPermissionGroupDisplayNames = UserCompareInit(firstUser);
            ICollection<string> secondUserPermissionGroupDisplayNames = UserCompareInit(secondUser);

            // Alphabetically sorting for better read experience
            List<string> fistList = firstUserPermissionGroupDisplayNames.OrderBy(x => x).ToList();
            List<string> sectList = secondUserPermissionGroupDisplayNames.OrderBy(x => x).ToList();

            // Lists to the permissions which, Both have and for the only one has
            ICollection<string> BelongsToBothList = new List<string>();
            ICollection<string> BelongsToOnlyOneList = new List<string>();

            // Permissions which belongs to both of them - with order by
            var BelongsToBoth = firstUserPermissionGroupDisplayNames.Intersect(secondUserPermissionGroupDisplayNames)
                .Select(a => new
                {
                    PermissionGroup = a,
                })
                .OrderBy(x => x.PermissionGroup);

            // Add this information to a list to store and later write into a file
            foreach (var d in BelongsToBoth)
            {
                BelongsToBothList.Add(d.PermissionGroup);
            }

            // Permissions which belongs to only one of them - with order by
            var BelongsToOnlyOne = firstUserPermissionGroupDisplayNames.Concat(secondUserPermissionGroupDisplayNames)
                .Except(firstUserPermissionGroupDisplayNames.Intersect(secondUserPermissionGroupDisplayNames)) 
                .Select(a => new
                {
                    PermissionGroup = a,
                    User = firstUserPermissionGroupDisplayNames.Any(c => c == a) ? firstUser : secondUser
                })
                .OrderBy(x => x.PermissionGroup);

            // Add this information to a list to store and later write into a file
            foreach (var d in BelongsToOnlyOne)
            {
                BelongsToOnlyOneList.Add(d.PermissionGroup + " only belongs to " + d.User);
            }

            // Create a file and write the information parts into it with append possibility
            SaveToTxt(fistList, firstUser, _connectionAdapter.ConnUri.ToString());
            SaveToTxt(sectList, secondUser);
            SaveToTxt(BelongsToBothList, "Both    ");
            SaveToTxt(BelongsToOnlyOneList, "OnlyOne ");

            // Create a MS Word file with color coded permission lists
            SaveToWord(fistList, sectList, BelongsToBothList, BelongsToOnlyOneList, firstUser, secondUser, _connectionAdapter.ConnUri.ToString());
        }

        /// <summary>
        /// Init a Collection from the user permission group DisplayNames
        /// </summary>
        /// <param name="userName">Name of the user to init permission group DisplayNames collection</param>
        /// <returns></returns>
        private ICollection<string> UserCompareInit(string userName)
        {
            // Get the indentities from the specific user Identity
            TeamFoundationIdentity[] userMemberOfIdentities = GetUserIdentityWithMemberOfs(userName);
            // Build a collection from the permission group DisplayNames by the 'Member of' identities
            ICollection<string> groupDisplayNames = BuildListFromPermissionDisplayNames(userMemberOfIdentities);

            if (groupDisplayNames != null && groupDisplayNames.Count != 0)
            {
                _Logger.Info("User: " + userName + " permission group DisplayNames init was successful");
                _Logger.Flush();

                return groupDisplayNames;
            }
            return null;
        }

        /// <summary>
        /// Get a Collection from the permission group DisplayNames
        /// </summary>
        /// <param name="userMemberOfIdentities">Member of identities</param>
        /// <returns>null or collection of the permission group DisplayNames</returns>
        private ICollection<string> BuildListFromPermissionDisplayNames(TeamFoundationIdentity[] userMemberOfIdentities)
        {
            // Array check
            if (userMemberOfIdentities == null || userMemberOfIdentities.Length == 0)
            {
                _Logger.Info("Empty collection (identities), cannot build a collection from the permission group DisplayNames");
                _Logger.Flush();

                return null;
            }

            // Result list with the group names from the user 'Member of' section
            ICollection<string> userMemberships = new List<string>();

            foreach (TeamFoundationIdentity identity in userMemberOfIdentities)
            {
                userMemberships.Add(identity.DisplayName);
            }

            return userMemberships;
        }

        /// <summary>
        /// Get the specific user 'Member of' identities from the complete server
        /// </summary>
        /// <param name="userName">Name of the user to get the 'Member of' identities</param>
        /// <returns>null or TeamFoundationIdentity array</returns>
        private TeamFoundationIdentity[] GetUserIdentityWithMemberOfs(string userName)
        {
            try
            {
                // Get the user identity by username
                TeamFoundationIdentity userIdentity = _connectionAdapter.identityManagementService.ReadIdentity(IdentitySearchFactor.AccountName,
                                                            userName,
                                                            MembershipQuery.Expanded,
                                                            ReadIdentityOptions.ExtendedProperties);

                // User's 'Member of' identities
                TeamFoundationIdentity[] memberOfIdentities = _connectionAdapter.identityManagementService.ReadIdentities(userIdentity.MemberOf,
                                                                    MembershipQuery.None, ReadIdentityOptions.None);

                if(memberOfIdentities != null && memberOfIdentities.Length != 0)
                {
                    _Logger.Info("User: " + userName + " identity info processed with successful");
                    _Logger.Flush();

                    return memberOfIdentities;
                }
                else
                {
                    _Logger.Info("User: " + userName + " has no 'Member of' identities");
                    _Logger.Flush();
                }
            }
            catch (Exception ex)
            {
                _Logger.Error("Problem: User identity problem with user: " + userName + " ,check the user account name. Extendend error message: " + ex.Message);
                _Logger.Flush();
                Environment.Exit(-1);
            }

            return null;
        }

        /// <summary>
        /// Create or append to a text file
        /// Write the Permission groups (DisplayNames) into the file with categorization by the owner(s)
        /// </summary>
        /// <param name="permissionGroups">Permission DisplayName group collection</param>
        /// <param name="belongsTo">This permissions belongs to ...</param>
        private void SaveToTxt(ICollection<string> permissionGroups, string belongsTo, string serverName = null)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string DateTitle = DateTime.Now.ToString();
            string dateCorrectForm = DateTitle.Replace("/", "-").Replace(":", "_");

            try
            {
                using (FileStream fileStream = new FileStream(path + @"\PermissionGroups " + dateCorrectForm + ".txt", FileMode.Append))
                using (TextWriter tw = new StreamWriter(fileStream))
                {
                    if (!string.IsNullOrEmpty(serverName))
                    {
                        tw.WriteLine("Server: " + serverName);
                    }

                    tw.WriteLine("Permissions of: " + belongsTo + " --------------------------------------------------------------------");
                    foreach (string group in permissionGroups)
                    {
                        tw.WriteLine(group);
                    }
                    tw.WriteLine("---------------------------------------------------------------------------------------------");
                    tw.Flush();
                    tw.Close();
                    fileStream.Close();
                    _Logger.Info("Permission detais write to the file was successful (Permission category was " + belongsTo + " )");
                    _Logger.Flush();
                }
            }
            catch (Exception ex)
            {
                _Logger.Error("Problem: " + ex.Message);
                _Logger.Flush();
            }
        }

        /// <summary>
        /// Create a word document and write into the users permissions with color code
        /// Green -> both have this permission
        /// Red -> only one has this permission
        /// </summary>
        /// <param name="FirstUserPermissionGroups">First user permissions</param>
        /// <param name="SecondUserPermissionGroups">Second user permissions</param>
        /// <param name="BelongsToBothList">Permission collection which both have</param>
        /// <param name="BelongsToOnlyOneList">Permission collection which only one has</param>
        /// <param name="firstUser">First user name - account name</param>
        /// <param name="secondUser">Second user name - account name</param>
        /// <param name="serverName">Connected server name</param>
        private void SaveToWord(ICollection<string> FirstUserPermissionGroups, ICollection<string> SecondUserPermissionGroups,
            ICollection<string> BelongsToBothList, ICollection<string> BelongsToOnlyOneList, 
            string firstUser, string secondUser, string serverName)
        {
            try
            {
                string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string DateTitle = DateTime.Now.ToString();
                string dateCorrectForm = DateTitle.Replace("/", "-").Replace(":", "_");

                // Creates an instance of WordDocument Instance (Empty Word Document)
                WordDocument document = new WordDocument();

                // Adds a new section into the Word document
                IWSection section = document.AddSection();

                // Adds paragraphs into the section
                IWParagraph firstParagraph = section.AddParagraph();
                IWParagraph secondParagraph = section.AddParagraph();
                IWParagraph thirdParagraph = section.AddParagraph();

                // Sets the paragraph's horizontal alignment as justify
                firstParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
                secondParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
                thirdParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;

                // Adds a text range into the paragraph
                IWTextRange firstTextRange = firstParagraph.AppendText("Server: " + serverName + '\n');
                firstTextRange.CharacterFormat.FontName = "Calibri";
                firstTextRange.CharacterFormat.FontSize = 16;

                IWTextRange bothHas = firstParagraph.AppendText("Both users have this permission \n");
                bothHas.CharacterFormat.FontName = "Calibri";
                bothHas.CharacterFormat.FontSize = 14;
                bothHas.CharacterFormat.TextColor = Color.Green;

                IWTextRange onlyOne = firstParagraph.AppendText("Only one user has this permission \n");
                onlyOne.CharacterFormat.FontName = "Calibri";
                onlyOne.CharacterFormat.FontSize = 14;
                onlyOne.CharacterFormat.TextColor = Color.Red;


                IWTextRange secondTextRange = secondParagraph.AppendText("Permissions of: " + firstUser +
                    " -------------------------------------------------------------------- \n");
                secondTextRange.CharacterFormat.FontName = "Calibri";
                secondTextRange.CharacterFormat.FontSize = 14;

                foreach (string group in FirstUserPermissionGroups)
                {
                    // Check the permission both have or not
                    if (BelongsToBothList.Contains(group))
                    {
                        // set and use the color settings for this
                        bothHas = secondParagraph.AppendText(group + '\n');
                        bothHas.CharacterFormat.FontName = "Calibri";
                        bothHas.CharacterFormat.FontSize = 14;
                        bothHas.CharacterFormat.TextColor = Color.Green;
                    }
                    // Get the permission by the group and check this permissiion belongs to this user only or not
                    // BelongsToOnlyOneList elements looks like this:  'PermissionGroup + " only belongs to " + User'
                    else if (BelongsToOnlyOneList.Where(x => x.Split(' ').First() == group).Where(x => x.Split(' ').Last() == firstUser) != null)
                    {
                        // set and use the color settings for this
                        onlyOne = secondParagraph.AppendText(group + '\n');
                        onlyOne.CharacterFormat.FontName = "Calibri";
                        onlyOne.CharacterFormat.FontSize = 14;
                        onlyOne.CharacterFormat.TextColor = Color.Red;
                    }
                }
                secondParagraph.AppendText("--------------------------------------------------------------------------------------------- \n");

                IWTextRange thirdTextRange = thirdParagraph.AppendText("Permissions of: " + secondUser +
                    " -------------------------------------------------------------------- \n");
                thirdTextRange.CharacterFormat.FontName = "Calibri";
                thirdTextRange.CharacterFormat.FontSize = 14;

                foreach (string group in SecondUserPermissionGroups)
                {
                    // Check the permission both have or not
                    if (BelongsToBothList.Contains(group))
                    {
                        // set and use the color settings for this
                        bothHas = thirdParagraph.AppendText(group + '\n');
                        bothHas.CharacterFormat.FontName = "Calibri";
                        bothHas.CharacterFormat.FontSize = 14;
                        bothHas.CharacterFormat.TextColor = Color.Green;
                    }
                    // Get the permission by the group and check this permissiion belongs to this user only or not
                    // BelongsToOnlyOneList elements looks like this:  'PermissionGroup + " only belongs to " + User'
                    else if (BelongsToOnlyOneList.Where(x => x.Split(' ').First() == group).Where(x => x.Split(' ').Last() == secondUser) != null)
                    {
                        // set and use the color settings for this
                        onlyOne = thirdParagraph.AppendText(group + '\n');
                        onlyOne.CharacterFormat.FontName = "Calibri";
                        onlyOne.CharacterFormat.FontSize = 14;
                        onlyOne.CharacterFormat.TextColor = Color.Red;
                    }
                }
                thirdParagraph.AppendText("--------------------------------------------------------------------------------------------- \n");

                //Save and close the Word document
                document.Save(path + @"\PermissionGroups " + dateCorrectForm + ".docx");
                document.Close();
                _Logger.Info("Permission detais write to the MS Word file was successful");
                _Logger.Flush();
            }
            catch (Exception ex)
            {
                _Logger.Error("Problem: " + ex.Message);
                _Logger.Flush();
            }
        }
    }
}
