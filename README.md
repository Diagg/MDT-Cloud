# MDT-Cloud
Deploy Latest Windows 10 Build without pre staging WIM images into MDT.  
Everything is retrieved from the cloud !

Installation details can be found With here: [OSD-Couture.com](https://www.osd-couture.com/2021/01/mdt-deploy-from-cloud-grab-latest-build.html)

## Release History

- V 1.0 - Initiale release.
- V 2.0 - 7zip is replaced by [Wimlib](https://wimlib.net/), Prevent Installation script from running on Powershell 7, ISO retrieval is now done during ZTI-Gather.wsf processing to prevent Wizard from freezing/bugging when selecting Windows Build.