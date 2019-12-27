import { IInputs, IOutputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import { Calendar, EventInput } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
import 'moment'
import moment = require("moment");
type DataSet = ComponentFramework.PropertyTypes.DataSet;

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
			firstDay: 1
		});

		// _calendar.addEvent({ // this object will be "parsed" into an Event Object
		// 	title: 'Business Review', // a property!
		// 	start: '2019-12-27 16:00:00', // a property!
		// 	end: '2019-12-27', // a property! ** see important note below about 'end' **
		// 	className: "recurring"
		// });

		this._calendar.render();

		container.appendChild(this._divCalendar);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view


		if (!context.parameters.appointmentsDataSet.loading) {

			let appointmentsRecordSet = context.parameters.appointmentsDataSet;

			appointmentsRecordSet.sortedRecordIds.forEach(recordId => {
				this._calendar.addEvent({
					title: appointmentsRecordSet.records[recordId].getValue("Subject").toString(), // a property!
					start: moment(appointmentsRecordSet.records[recordId].getValue("Start Time").toString(), "DD/MM/YYYY H:mm").format("YYYY-MM-DD H:mm:00"), 
					end: moment(appointmentsRecordSet.records[recordId].getValue("End Time").toString(), "DD/MM/YYYY H:mm").format("YYYY-MM-DD H:mm:00"),
					className: appointmentsRecordSet.records[recordId].getValue("Appointment Type").toString().toLowerCase().replace(" ","-")
				});

				
				// // this._calendar.addEvent({ // this object will be "parsed" into an Event Object
				// // 	title: 'Business Review', // a property!
				// // 	start: '2019-12-27', // a property!
				// // 	end: '2019-12-27', // a property! ** see important note below about 'end' **
				// // 	className: "recurring"
				// // });
				// 	if (column.name == "tag") {
				// 		var tagDiv = <HTMLSpanElement>document.createElement("div");
				// 		tagDiv.className = "tags";
				// 		var tags = (<string>recordSet.records[recordId].getValue(column.name)).split(";");

				// 		tags.forEach(tag => {
				// 			var tagSpan = <HTMLSpanElement>document.createElement("span");
				// 			tagSpan.className = "tag";
				// 			tagSpan.innerText = tag;

				// 			if (tag == "urgent")
				// 				tagSpan.style.background = "#E51400";
				// 			if (tag == "blocked")
				// 				tagSpan.style.background = "#FF9642";
				// 			if (tag == "green")
				// 				tagSpan.style.background = "#668D3C";
				// 			if (tag == "stage gate")
				// 				tagSpan.style.background = "#007996";

				// 			tagDiv.appendChild(tagSpan);
				// 		});

				// 		recordDiv.appendChild(tagDiv);
				// 	}
				// 	else {
				// 		var span = <HTMLSpanElement>document.createElement("span");
				// 		span.className = "element";
				// 		span.innerText = <string>recordSet.records[recordId].getValue(column.name);
				// 		recordDiv.appendChild(span);
				// 	}

				// });
			});

			this._calendar.render();
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