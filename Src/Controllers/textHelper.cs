using System.Text;

namespace CreateXslt;

public class textHelper
{
    public static string GetLandingPageText()
    {
        StringBuilder sb = new StringBuilder();

        sb.AppendLine("\n Made by Karim \n");
        sb.AppendLine(" Her er et lille program til at lave crm rapporter for dansk metal.");
        sb.AppendLine(" Programmet bruger du på følgene måde:\n");
        sb.AppendLine("     1. Udvikle in sql og gem resultatet i en csv fil med kolonne overskift som den øverste line.\n");
        sb.AppendLine("     2. Åben programmet og indlæs csv filen.\n");
        sb.AppendLine("     3. Gennemgå hver kolonne du skal tjekke:\n");
        sb.AppendLine("         a. Kolonne navn, programmet forslår kolonne navnet hvor _ er erstattet med space og fjerner 'dm_' i starten.");
        sb.AppendLine("         b. Excel filter typen, programmet analysere selv et forslag til et filter.");
        sb.AppendLine("         c. CRM rapport input typen, programmet analysere selv et forslag til inputtype. (ikke supporteret endnu)");

        return sb.ToString();
    }
}