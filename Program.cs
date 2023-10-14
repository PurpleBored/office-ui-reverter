using Microsoft.Win32;
using System.Diagnostics;


class officeuireverter
{
    static void AddRegistryKeysForOfficeApp(string officeApp)
    {
        string keyPath = $@"SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\{officeApp}";

        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath))
        {
            if (key != null)
            {
                key.SetValue("Microsoft.Office.UXPlatform.FluentSVRefresh", "false", RegistryValueKind.String);
                key.SetValue("Microsoft.Office.UXPlatform.FluentSVRibbonOptionsMenu", "false", RegistryValueKind.String);
                key.SetValue("Microsoft.Office.UXPlatform.RibbonTouchOptimization", "false", RegistryValueKind.String);

                    Console.WriteLine($"Registry keys added for {officeApp}.");
            }
            else
            {
                Console.WriteLine($"Error creating/opening registry key for {officeApp}.");
            }
        }
    }

    static void RemoveRegistryKeysForOfficeApp(string officeApp)
    {
        string keyPath = $@"SOFTWARE\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\{officeApp}";

        try
        {
            Registry.CurrentUser.DeleteSubKeyTree(keyPath);
            Console.WriteLine($"Registry keys removed for {officeApp}.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error removing registry keys for {officeApp}: {ex.Message}");
        }
    }
    static string ascii_logo = @"
       ___   __  __ _             _   _ ___    ____                     _            
      / _ \ / _|/ _(_) ___ ___   | | | |_ _|  |  _ \ _____   _____ _ __| |_ ___ _ __ 
     | | | | |_| |_| |/ __/ _ \  | | | || |   | |_) / _ \ \ / / _ \ '__| __/ _ \ '__|
     | |_| |  _|  _| | (_|  __/  | |_| || |   |  _ <  __/\ V /  __/ |  | ||  __/ |   
      \___/|_| |_| |_|\___\___|   \___/|___|  |_| \_\___| \_/ \___|_|   \__\___|_|   
                                                                               
    ";
    static void ShowMenu()
    {
        Console.WriteLine(ascii_logo);
        Console.WriteLine("Menu:");
        Console.WriteLine("1. Revert back to old UI");
        Console.WriteLine("2. Revert back to new UI");
        Console.WriteLine("3. Exit");
        Console.Write("Enter your choice (1-3): ");
    }
    
    static void Main()
    {
        bool exit = false;

        while (!exit)
        {
            ShowMenu();
            string choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    string[] officeAppsToAdd = new string[]
                    {
                        "access", "excel", "onenote", "outlook", "powerpoint", "publisher", "visio", "word"
                    };

                    foreach (string app in officeAppsToAdd)
                    {
                        AddRegistryKeysForOfficeApp(app);
                    }
                    break;
                case "2":
                    string[] officeAppsToRemove = new string[]
                    {
                        "access", "excel", "onenote", "outlook", "powerpoint", "publisher", "visio", "word"
                    };

                    foreach (string app in officeAppsToRemove)
                    {
                        RemoveRegistryKeysForOfficeApp(app);
                    }
                    break;
                case "3":
                    exit = true;
                    break;
                default:
                    Console.WriteLine("Invalid choice. Please try again.");
                    break;
            }
        }
    }
}
