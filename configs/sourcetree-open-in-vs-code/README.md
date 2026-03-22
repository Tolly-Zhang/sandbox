# Open Repository in VS Code from SourceTree

This configuration is for a custom action inside SourceTree that allows you to open VS Code from the repo root.

## Setup

1. Open SourceTree and navigate to Tools > Options > Custom Actions.
2. Create a custom action with the given configuration.

## Configuration

- - [ ] Open in a separate window
- - [ ] Show full output
- - [x] Run command silently
- **Script to run**: `C:\Users\*\AppData\Local\Programs\Microsoft VS Code\bin\code.cmd`
- **Parameters**: `-r "$REPO"`

## Usage

In Sourcetree, go to **Actions** > **Custom Actions** then run.