//import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
//import { Dialog } from '@microsoft/sp-dialog';

import { escape } from '@microsoft/sp-lodash-subset';
import  styles from "./AppCustomizerPlaceholdersApplicationCustomizer.module.scss";

import { IMission } from "../../models";
import { MissionService } from "../../services";

//import * as strings from 'AppCustomizerPlaceholdersApplicationCustomizerStrings';

//const LOG_SOURCE: string = 'AppCustomizerPlaceholdersApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppCustomizerPlaceholdersApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppCustomizerPlaceholdersApplicationCustomizer
  extends BaseApplicationCustomizer<IAppCustomizerPlaceholdersApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;


  public onInit(): Promise<void> {
    //Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // let message: string = this.properties.testMessage;
    // if (!message) {
    //   message = '(No properties were provided.)';
    // }

    // Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`).catch(() => {
    //   /* handle error */
    // });

    this._renderPlaceholders();

    return Promise.resolve();
  }

  private _onDispose(): void { }

  private _renderPlaceholders(): void {

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        {
          onDispose: this._onDispose
        }
      );

      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) wasn\'t found.');
        return;
      }
    };

    if (this._topPlaceholder.domElement) {
      this._topPlaceholder.domElement.innerHTML = this._getPlaceholderHtml(MissionService.getMission('AS-506'), 'Moon Landing');
    }

    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        {
          onDispose: this._onDispose
        }
      );

      if (!this._bottomPlaceholder) {
        console.error('The expected placeholder (Top) wasn\'t found.');
        return;
      }
    };

    if (this._bottomPlaceholder.domElement) {
      this._bottomPlaceholder.domElement.innerHTML = this._getPlaceholderHtml(MissionService.getMission('AS-512'), 'last moon visit');
    }

  }

  private _getPlaceholderHtml(mission: IMission, prefixMessage: string){
    const missionTime: string = `${this._getLocalizedTimeString(new Date(mission.launch_date))}`;
    const placeHolderBody: string = `
      <div class="${styles.app}">
        ${escape(prefixMessage)}: ${escape(mission.name)} on ${escape(missionTime)}
      </div>
    `;
    return placeHolderBody;
  }

  private _getLocalizedTimeString(dateTimestamp: Date): string {
    return `${this._getMonthName(dateTimestamp.getMonth())} ${dateTimestamp.getDate()}, ${dateTimestamp.getFullYear()}`;
  }

  private _getMonthName (monthIndex: number): string{
    const monthNames: string[] = [
      'January',
      'February',
      'March',
      'April',
      'May',
      'June',
      'July',
      'August',
      'September',
      'October',
      'November',
      'December'
    ];
    return monthNames[monthIndex];
  }
}