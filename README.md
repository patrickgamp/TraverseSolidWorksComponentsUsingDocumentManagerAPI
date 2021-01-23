# TraverseSolidWorksComponentsUsingDocumentManagerAPI
Traverse all components in a SolidWorks model using the Document Manager API

Give a SolidWorks file (most likely a top level assembly), we can have multiple tasks running to get all components' information recursively.
Each task handles a single component.
If hit a same named component, we will search in all running tasks, and wait for the task to end, then clone the item, instead of getting the info again using Document Manager API.

There should be a static global string array indicating the custom properties in the SolidWorks document：
public static string[] PropertyNamesOftenUsed = new string[]
        {
            "DrawNo",
            "PARTNAME"
        };
You should replace with your own definitions
