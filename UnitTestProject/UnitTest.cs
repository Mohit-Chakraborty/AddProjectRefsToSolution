using Microsoft.VisualStudio.Sdk.TestFramework;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VSIXProject;

namespace UnitTestProject
{
    [TestClass]
    public class UnitTest
    {
        internal static GlobalServiceProvider MockServiceProvider { get; private set; }

        [AssemblyInitialize]
        public static void AssemblyInit(TestContext context)
        {
            MockServiceProvider = new GlobalServiceProvider();
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            MockServiceProvider.Dispose();
        }

        [TestInitialize]
        public void TestInit()
        {
            MockServiceProvider.Reset();
        }

        [TestMethod]
        public async System.Threading.Tasks.Task TestMethod1Async()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var mockVsHierarchy = new MockVsHierarchy();

            var originalPath = @"C:\foo\bar.csproj";
            var expectedPath = originalPath;

            var resolvedPath = ProjectHelper.ResolveMacrosInPath(mockVsHierarchy, originalPath);
            Assert.AreEqual<string>(expectedPath, resolvedPath);
        }

        [TestMethod]
        public async System.Threading.Tasks.Task TestMethod2()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var mockVsHierarchy = new MockVsHierarchy();

            var originalPath = @"C:\$(foo)\bar.csproj";
            var expectedPath = @"C:\foo\bar.csproj";

            var resolvedPath = ProjectHelper.ResolveMacrosInPath(mockVsHierarchy, originalPath);
            Assert.AreEqual<string>(expectedPath, resolvedPath);
        }

        [TestMethod]
        public async System.Threading.Tasks.Task TestMethod3()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var mockVsHierarchy = new MockVsHierarchy();

            var originalPath = @"C:\$(foo)\$(bar).csproj";
            var expectedPath = @"C:\foo\bar.csproj";

            var resolvedPath = ProjectHelper.ResolveMacrosInPath(mockVsHierarchy, originalPath);
            Assert.AreEqual<string>(expectedPath, resolvedPath);
        }

        [TestMethod]
        public async System.Threading.Tasks.Task TestMethod4()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var mockVsHierarchy = new MockVsHierarchy();

            var originalPath = @"C:\$(foo)\$(bar.csproj)";
            var expectedPath = @"C:\foo\bar.csproj";

            var resolvedPath = ProjectHelper.ResolveMacrosInPath(mockVsHierarchy, originalPath);
            Assert.AreEqual<string>(expectedPath, resolvedPath);
        }
    }
}
