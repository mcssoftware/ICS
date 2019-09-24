export default class IcsAppConstants {
    public static getCommitteesPartial: () => string = () => { return "/api/Committees/"; };
    public static getCommitteesStaffPartial: () => string = () => { return "api/Calendar/Committee"; };
    public static getazureServiceUrl: () => string = () => { return "https://WYOLEG.GOV/LsoOffice365Service"; };

    public static getMaterialPreviewDocPartial: () => string = () => { return "api/Committee/CreateMaterialIndex?preview=true"; };
    public static getCreateMeetingNoticePartial: () => string = () => { return "api/Committee/CreateMeetingNotice"; };
    public static getCreateAgendaPreviewPartial: () => string = () => { return "api/Committee/CreateAgenda"; };
    public static getCreateMinutePreviewPartial: () => string = () => { return "api/Committee/CreateMinutes"; };

    public static getPreviewFolder: () => string = () => { return "Preview"; };
}