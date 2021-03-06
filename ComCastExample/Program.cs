using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using mshtml;
using SHDocVw;

namespace ComCastExample
{
    internal class Program
    {
        public static async Task<int> Main(string[] args)
        {
            var ie = await StartInternetExplorer(TimeSpan.FromSeconds(10));
            Console.WriteLine($"Started {ie.Name}");

            await UrlLoaded(ie);
            Console.WriteLine($"Navigation complete to {ie.LocationURL}");

            var document = ie.Document as HTMLDocument;

            if (document == null)
                return 1;

            var element = document.getElementById("fname");
            var isHtmlFrameElement = element is DispHTMLInputElement;
            Console.WriteLine($"Element has Tag {element.tagName} and value {element.getAttribute("value")}");
            Console.WriteLine($"Element is {(element is DispHTMLInputElement ? "" : "NOT ")}a DispHTMLInputElement element");
            Console.WriteLine($"Element is {(element is HTMLInputElementClass ? "" : "NOT ")}a HTMLInputElementClass element");
            Console.WriteLine($"Element is {(element is HTMLInputElement ? "" : "NOT ")}a HTMLInputElement element");
            Console.WriteLine($"Element is {(element is IHTMLInputElement ? "" : "NOT ")}a IHTMLInputElement element");

            Console.WriteLine($"Element is {(element is DispHTMLFrameElement ? "" : "NOT ")}a DispHTMLFrameElement element");
            Console.WriteLine($"Element is {(element is HTMLFrameElementClass ? "" : "NOT ")}a HTMLFrameElementClass element");
            Console.WriteLine($"Element is {(element is HTMLFrameElement ? "" : "NOT ")}a HTMLFrameElement element");
            Console.WriteLine($"Element is {(element is IHTMLFrameElement ? "" : "NOT ")}a IHTMLFrameElement element");
            return 0;
        }

        private static async Task UrlLoaded(InternetExplorer ie)
        {
            var tcs = new TaskCompletionSource<bool>();
            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            cts.Token.Register(() => tcs.TrySetCanceled());
            ie.NavigateComplete2 += (object _, ref object _) => tcs.TrySetResult(true);
            if (!ie.Busy)
                tcs.TrySetResult(true);

            await tcs.Task;
        }

        private static async Task<InternetExplorer> StartInternetExplorer(TimeSpan timeout)
        {
            var shellWindows = new ShellWindows();
            var tcs = new TaskCompletionSource<InternetExplorer>();
            var cts = new CancellationTokenSource(timeout);
            cts.Token.Register(() => tcs.TrySetCanceled());
            var ignoreList = new HashSet<InternetExplorer>(shellWindows.OfType<InternetExplorer>());
            var url = "https://www.w3schools.com/html/html_forms.asp";
            shellWindows.WindowRegistered += _ =>
            {
                foreach (var window in shellWindows
                             .OfType<InternetExplorer>()
                             .Except(ignoreList)
                             .Where(ie => ie.LocationURL == url))
                    tcs.TrySetResult(window);
            };

            Process.Start("IExplore.exe", url);
            return await tcs.Task;
        }
    }
}