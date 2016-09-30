
export class Utils {
  public static getUserPhotoUrl(userEmail: string, siteUrl: string, size: string = 'S'): string {
    return `${siteUrl}/_layouts/15/userphoto.aspx?size=${size}&accountname=${userEmail}`;
  }
  public static trim(s: string): string {
    if (s && s.length > 0) {
      return s.replace(/^\s+|\s+$/gm, '');
    }
    else {
      return s;
    }
  }

  public static getInitialsFromFullName(fullName: string): string {
    var names = fullName.split(' ');
    return names[0].substring(0,1)+names[1].substring(0,1)
  }
}