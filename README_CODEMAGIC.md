# Music Label Manager - Codemagic Ready

## Included
- codemagic.yaml for Android release APK
- Flutter source files
- Codemagic build instructions

## Important: Android platform folder
If this project does not contain an `android/` folder, generate the Flutter platform files before building:

flutter create .

Then upload the complete project to GitHub.

## Codemagic steps
1. Create a GitHub repository.
2. Upload all files from this project.
3. Open Codemagic and connect the GitHub repository.
4. Select the `android-release` workflow.
5. Start build.
6. Download `app-release.apk` from Build Artifacts.

Build command:
flutter build apk --release
