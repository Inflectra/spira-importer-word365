## Setting up dev enviroment

## Setting up dev enviroment

- **1** -- Refer to the page https://docs.microsoft.com/en-us/office/dev/add-ins/overview/learning-path-beginner
- **2** -- In step 2 choose node JS and Visual Studio Code
-- **2.1** -- If you don't have Node.js and npm installed, follow the instructions from and install them https://docs.microsoft.com/en-us/office/dev/add-ins/overview/set-up-your-dev-environment
-- **2.2** -- Run the cmd command: `npm install -g yo generator-office`
- **3** -- Open the root folder in VS Code, click Ctrl + Shift + ` to open the terminal - inside VS Code, and run the command 'npm install'
- **4** -- Run the commands `npm run build` then  `npm run start` to run your project.
- **5** -- There will be a pop up asking you to validate a localhost security certificate, do validate the certificate or it will not work.

## Troubleshooting

If you have an existing node.js install with npm version < 7 you will get a lockfileversion error, to resolve this, delete package-lock.json and run `npm install` again, then follow set-up steps 4 and 5 again.