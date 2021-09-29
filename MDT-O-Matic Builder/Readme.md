# MDT-O-MATIC Builder

This set of tool allow to build the final script (**MDT-O-Matic.Ps1**).

## Process

Using **package builder.ps1** with the -DSSource argument referring to the path of the local deployment share used to create the solution.

The following actions will occur:
- files from folder **EmbeddedContent** will be packed to an archive.
- The archive will be embedded in the **MDT-O-Matic-Template-v5.ps1** file to create the final **MDT-O-MATIC.ps1** file in the parent folder.