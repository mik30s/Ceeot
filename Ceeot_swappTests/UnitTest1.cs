using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Ceeot_swapp;

namespace Ceeot_swappTests
{
    [TestClass]
    public class ProjectManagerTest
    {
        [TestMethod]
        public void createProject_All_Empty()
        {
            // Arrange

            // Act
            var projectManager = new ProjectManager();
            projectManager.createProject("","","","", SwattProject.ProjectVersion.APEX_0604, SwattProject.ProjectVersion.SWATT_2005);
            // Assert
            // use mocking framework here.
        }
    }
}
