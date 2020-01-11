import { IInputs, IOutputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import { Calendar, EventInput } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
// import {Tooltip} from 'tooltip-js'
const Tooltip = require('tooltip-js');
import moment = require("moment");
import * as $ from 'jquery';

export class CalendarView implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _divCalendar: HTMLDivElement;
	private _calendar: Calendar;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		// Add control initialization code
		this._divCalendar = document.createElement("div");
		this._divCalendar.id = "calendar";

		this._calendar = new Calendar(this._divCalendar, {
			plugins: [dayGridPlugin],
			defaultView: 'dayGridMonth',
			firstDay: 1,
			eventRender: function (info) {
				info.el.setAttribute("title", info.event.extendedProps["description"]);
			},
			contentHeight: 530
		});

		this._calendar.render();

		container.appendChild(this._divCalendar);

		//TODO define the schema name for the different attributes from the input
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
		if (!context.parameters.appointmentsDataSet.loading) {

			let appointmentsRecordSet = context.parameters.appointmentsDataSet;

			this._calendar.removeAllEvents();
			appointmentsRecordSet.sortedRecordIds.forEach(recordId => {
				this._calendar.addEvent({
					title: appointmentsRecordSet.records[recordId].getValue("subject").toString(), // a property!
					start: moment(appointmentsRecordSet.records[recordId].getValue("scheduledstart").toString(), "YYYY-MM-DDTHH:mm:ss.SSSZ").format("YYYY-MM-DD HH:mm:00"),
					end: moment(appointmentsRecordSet.records[recordId].getValue("scheduledend").toString(), "YYYY-MM-DDTHH:mm:ss.SSSZ").format("YYYY-MM-DD HH:mm:00"),
					className: this.mapValueToText(appointmentsRecordSet.records[recordId].getValue("new_appointmenttype") != null ? appointmentsRecordSet.records[recordId].getValue("new_appointmenttype").toString().toLowerCase().replace(" ", "-") : ""),
					description: appointmentsRecordSet.records[recordId].getValue("description") != null ? appointmentsRecordSet.records[recordId].getValue("description").toString() : ""
				});
			});

			this._calendar.render();
		}

	}


	//todo remove the hard coding and turn into reusable component
	private mapValueToText(value: string): string {
		if (value === "100000001") {
			return "stage-gate";
		}
		else if (value === "100000002") {
			return "urgent";
		}
		else {
			return "recurring";
		}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}

}