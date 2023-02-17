import { IUser } from "../interfaces/types/IUser";

export class User implements IUser {
    public Email: string = "";
    public Id: number = null;
    public LoginName: string = "";
    public Title: string = "";
    public FirstName: string = "";
    public LastName: string = "";
}