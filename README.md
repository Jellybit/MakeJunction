# MakeJunction
A Powershell script with GUI to simplify and speed up creating junction folders.

![image](https://github.com/user-attachments/assets/f1ae8519-8b1a-4340-bdc4-429296b54868)

I use a lot of linked folders so that I can store certain files on fast SSDs, and access them in different projects. This is done using a Windows feature called **Directory Junctions**. Junctions basically make Windows see a folder anywhere you want it to be seen, like a shortcut, but it pretends to actually be that folder. For a while, I was using Windows' built-in "mklink" command, but typing that command by hand took a while and got annoying. So I went and made this GUI to handle a lot of use cases, and even batch a ton of them at once.

## Features

* **Flexible Creation Options:**
    * **Make In Folder Tab:** Choose your original folder (Content Path) and a place where you want the link to appear (Link Path). A linked folder with the same name as your original folder will be created inside the Link Path.
    * **Replace Folder Tab:** Choose your original folder (Content Path) and an *existing* folder (Link Path) that you want to turn into a direct link to your content. This is useful if you want to free up space by moving a folder's contents and replacing the original with a link.
    * **From Text Tab:**
        * Automatically logs junctions you create using the other tabs.
        * Lets you paste or type a list of multiple junctions to create them in a batch.
* **Smart Path Input:**
    * **Browse Buttons:** Easily find the folders you want to use.
    * **Drag & Drop:** Drag folders directly from File Explorer into the path boxes.
* **Safety First:**
    * Checks if a folder or link with the same name already exists.
    * Asks for confirmation before replacing existing links or folders that contain files.
    * Automatically replaces *empty* folders on the "Replace Folder" tab to streamline the process.

## How to Use

1.  **Run the Script:** Simply run the `MakeJunctions.ps1`.
2.  The main window will appear with three tabs.

### Tab 1: Make In Folder

This tab is for when you want to create a linked folder *inside* another folder.

* **Content Path:** Select or drag your original folder here (e.g., `D:\MyActualGames`).
* **Link Path:** Select or drag the folder where you want the link to *appear within* (e.g., `C:\Program Files\Games`).
* **Action:** If your Content Path was `D:\MyActualGames`, this will create a linked folder at `C:\Program Files\Games\MyActualGames` that points to `D:\MyActualGames`.
* Click **"Make Junction"**.

### Tab 2: Replace Folder

This tab is for when you want an *existing folder path* to become a link to your content folder.

* **Content Path:** Select or drag your original folder here (e.g., `D:\MyMovedPhotos`).
* **Link Path:** Select or drag the folder path that you want to *become* the link (e.g., `C:\Users\YourName\Pictures\OldPhotos`). This exact path will be turned into a link.
* **Action:** The folder `C:\Users\YourName\Pictures\OldPhotos` will be replaced by a link pointing to `D:\MyMovedPhotos`.
    * If `C:\Users\YourName\Pictures\OldPhotos` was empty, it will be replaced automatically.
    * If it contained files or was an existing link, you'll be asked for confirmation.
* Click **"Replace Folder with Junction"**.

### Tab 3: From Text

![image](https://github.com/user-attachments/assets/7fd6d223-4fd4-4656-ab98-53749f16e4a7)

This tab is for creating multiple junctions from a list and for seeing a history of junctions created in the current session.

* **Instructions:** "Add a list of junctions to be made here. Format is Content Path and Link Path, separated by a tab. Any junctions made in the other tabs will be added here for future reference."
* **How it works:**
    1.  Each line in the big text box should have:
        `Path\To\Your\Original\ContentFolder` (then press the **Tab key**) `Path\Where\You\Want\The\Link`
    2.  Junctions you create using the "Make In Folder" or "Replace Folder" tabs will automatically be added to this list.
    3.  You can also type or paste your own list here.
    4.  Click **"Make Junctions from List"**.
        * The tool will try to create each link.
        * It will automatically replace empty folders at the link path.
        * It will *skip* creating a link if the link path already exists and isn't an empty folder (e.g., it has content, is already a link, or is a file). No pop-ups will appear for these skips.
        * You'll get a single summary message at the end.

## Important Notes

* **Use Local Paths:** For now, this tool is designed to work with folders on your local computer's drives (like `C:\`, `D:\`, etc.). Using network paths (like `\\server\share`) for either the Content Path or the Link Path is not supported in this version and will result in an error.
* **Permissions (Admin Rights or Developer Mode):** Creating junctions is a special operation.
    * If you encounter errors about "insufficient privilege," you might need to run the script **as an Administrator**.
    * Alternatively, on Windows 10/11, enabling **Developer Mode** in your Windows Settings usually allows you to create junctions without running as an administrator, if permissions are an issue.

## Compile Standalone

I took these steps to compile a standalone version that doesn't have the console window. Open Powershell and type the following commands:

* Install-Module ps2exe
* Set-ExecutionPolicy RemoteSigned -Scope Process -Force
* ps2exe "MakeJunctions.ps1" "MakeJunctions.exe" -noConsole

This installs ps2exe, which compiles powershell scripts. The execution policy command changes permissions to be able to compile this executable, but the permissions will revert to normal after you close this window.

---

Hope this tool is useful to someone other than me.
