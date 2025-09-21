using Microsoft.Xrm.Sdk;
using MsCrmTools.UserRolesManager.AppCode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace MsCrmTools.UserRolesManager.UserControls
{
    public partial class PrincipalSelector : UserControl
    {
        private int currentColumnOrder;
        private IOrganizationService service;

        public PrincipalSelector()
        {
            InitializeComponent();

            cbbType.SelectedIndexChanged -= cbbType_SelectedIndexChanged;
            cbbType.SelectedIndex = 0;
            cbbType.SelectedIndexChanged += cbbType_SelectedIndexChanged;
        }

        public List<Entity> SelectedItems
        {
            get { return lvUsersAndTeams.SelectedItems.Cast<ListViewItem>().Select(e => (Entity)e.Tag).ToList(); }
        }

        public IOrganizationService Service
        {
            set
            {
                service = value;
                cbbType_SelectedIndexChanged(null, null);
            }
        }

        public void LoadViews()
        {
            cbbType_SelectedIndexChanged(null, null);
        }

        public void LoadUsersWithRole(Guid roleId)
        {
            if (service == null)
                throw new Exception("IOrganization service is not initialized");

            lvUsersAndTeams.Items.Clear();
            lblSelection.Text = "Loading users with selected role...";

            var query = new Microsoft.Xrm.Sdk.Query.QueryExpression("systemuser");
            query.ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("systemuserid", "firstname", "lastname", "businessunitid");
            query.Criteria.AddCondition("isdisabled", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, false);
            
            var roleLink = query.AddLink("systemuserroles", "systemuserid", "systemuserid");
            roleLink.AddCondition("roleid", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, roleId);

            var users = service.RetrieveMultiple(query);
            
            lvUsersAndTeams.Items.AddRange(users.Entities.Select(user => new ListViewItem
            {
                Text = user.GetAttributeValue<string>("lastname") ?? "",
                ImageIndex = 0,
                StateImageIndex = 0,
                Tag = user,
                SubItems =
                {
                    user.GetAttributeValue<string>("firstname") ?? "",
                    user.GetAttributeValue<Microsoft.Xrm.Sdk.EntityReference>("businessunitid")?.Name ?? ""
                }
            }).ToArray());

            lblSelection.Text = $"Users found: {users.Entities.Count} out of {users.Entities.Count}";
        }

        private void cbbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (service == null)
            {
                throw new Exception("IOrganization service is not initialized for this control");
            }

            var vManager = new ViewManager(service);
            var items = new List<ViewItem>();
            lvUsersAndTeams.Columns.Clear();

            switch (cbbType.SelectedIndex)
            {
                case 0:
                    {
                        items = vManager.RetrieveViews("systemuser");

                        lvUsersAndTeams.Columns.AddRange(new[]
                    {
                        new ColumnHeader{Text = "Last name", Width = 150},
                        new ColumnHeader{Text = "First name", Width = 150},
                        new ColumnHeader{Text = "Business unit", Width = 150}
                    });
                    }
                    break;

                case 1:
                    {
                        items = vManager.RetrieveViews("team");

                        lvUsersAndTeams.Columns.AddRange(new[]
                    {
                        new ColumnHeader{Text = "Name", Width = 150},
                        new ColumnHeader{Text = "Business unit", Width = 150}
                    });
                    }
                    break;
            }

            if (items != null)
            {
                cbbViews.SelectedIndexChanged -= cbbViews_SelectedIndexChanged;
                cbbViews.Items.Clear();
                cbbViews.Items.AddRange(items.ToArray());
                cbbViews.SelectedIndexChanged += cbbViews_SelectedIndexChanged;
                cbbViews.SelectedIndex = 0;
            }
        }

        private void cbbViews_SelectedIndexChanged(object sender, EventArgs e)
        {
            lvUsersAndTeams.Items.Clear();
            var entName = "Users";
            var viewItem = (ViewItem)cbbViews.SelectedItem;
            var entity = QueryHelper.GetItems(viewItem.FetchXml, service);

            if (entity.EntityName == "systemuser")
            {
                lvUsersAndTeams.Items.AddRange(entity.Entities.ToList().Select(record => new ListViewItem
                {
                    Text = record.GetAttributeValue<string>("lastname"),
                    ImageIndex = 0,
                    StateImageIndex = 0,
                    Tag = record,
                    SubItems =
                    {
                        record.GetAttributeValue<string>("firstname"),
                        record.GetAttributeValue<EntityReference>("businessunitid").Name
                    }
                }).ToArray());
            }
            else if (entity.EntityName == "team")
            {
                entName = "Teams";

                lvUsersAndTeams.Items.AddRange(entity.Entities.ToList().Select(record => new ListViewItem
                {
                    Text = record.GetAttributeValue<string>("name"),
                    ImageIndex = 1,
                    StateImageIndex = 1,
                    Tag = record,
                    SubItems =
                    {
                        record.GetAttributeValue<EntityReference>("businessunitid").Name
                    }
                }).ToArray());
            }

            // tell users how many teams/users in the view
            labelDetails.Text = $"{entName} total: {lvUsersAndTeams.Items.Count}";
            labelDetails.Tag = entName;
        }

        /// <summary>
        /// Update the selection count for users/teams out of total
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lvUsersAndTeams_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            labelDetails.Text = $"{labelDetails.Tag.ToString()}: {lvUsersAndTeams.SelectedItems.Count} selected of {lvUsersAndTeams.Items.Count} total";
        }

        private void lvUsersAndTeams_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (e.Column == currentColumnOrder)
            {
                lvUsersAndTeams.Sorting = lvUsersAndTeams.Sorting == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
                lvUsersAndTeams.ListViewItemSorter = new ListViewItemComparer(e.Column, lvUsersAndTeams.Sorting);
            }
            else
            {
                currentColumnOrder = e.Column;
                lvUsersAndTeams.Sorting = SortOrder.Ascending;
                lvUsersAndTeams.ListViewItemSorter = new ListViewItemComparer(e.Column, SortOrder.Ascending);
            }
        }

    }
}