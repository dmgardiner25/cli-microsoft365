import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate,
  CommandError
} from '../../../../Command';
import request from '../../../../request';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import config from '../../../../config';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  emails: string;
}

class SpoPersonalSiteAddCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PERSONALSITE_ADD}`;
  }

  public get description(): string {
    return 'Creates personal site for the specified users';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let spoAdminUrl: string = '';

    const emailObjects = this.createEmailObjects(args.options.emails.split(','));

    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((_spoAdminUrl: string): Promise<ContextInfo> => {
        spoAdminUrl = _spoAdminUrl;

        return this.getRequestDigest(spoAdminUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        let requestBody = '';

        if (this.verbose) {
          cmd.log(`Creating Personal Sites. Please wait, this might take a moment...`);
        }

        requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="5" ObjectPathId="4" /><Method Name="CreatePersonalSiteEnqueueBulk" Id="6" ObjectPathId="4"><Parameters><Parameter Type="Array">${emailObjects}</Parameter></Parameters></Method></Actions><ObjectPaths><StaticMethod Id="4" Name="GetProfileLoader" TypeId="{9c42543a-91b3-4902-b2fe-14ccdefb6e2b}" /></ObjectPaths></Request>`;

        console.log(requestBody);

        const requestOptions: any = {
          url: `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`,
          headers: {
            'X-RequestDigest': res.FormDigestValue
          },
          body: requestBody
        };

        return request.post(requestOptions);
      })
      .then((res: string): void => {
        const json: ClientSvcResponse = JSON.parse(res);
        const response: ClientSvcResponseContents = json[0];
        if (response.ErrorInfo) {
          cb(new CommandError(response.ErrorInfo.ErrorMessage));
          return;
        }
        else {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  private createEmailObjects(emails: string[]): string {
    return emails.map((email) => {
      return `<Object Type="String">${email}</Object>`
    }).join('');
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-e, --emails <emails>',
        description: 'Comma-separated list of e-mail addresses of users for whom to create a personal site'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.emails) {
        return 'Required parameter emails missing';
      }

      const emails = args.options.emails.split(',');
      const invalidEmails: string[] = [];
      emails.forEach((email: string) => {
        if (!Utils.isValidEmail(email)) { invalidEmails.push(email) };
      });

      if (invalidEmails.length !== 0) {
        return `${invalidEmails.join(', ')} not valid email`
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    If you choose to promote the page using the ${chalk.blue('promoteAs')} option
    or enable page comments, you will see the result only after publishing
    the page.

  Examples:

    Creates a new personal sites for multiple users
      ${this.name} --emails katiej@contoso.onmicrosoft.com,garth@contoso.onmicrosoft.com

`);
  }
}

module.exports = new SpoPersonalSiteAddCommand();