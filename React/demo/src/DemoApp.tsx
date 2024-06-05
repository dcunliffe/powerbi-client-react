// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useEffect } from 'react';
import { models, Report, Embed, service, Page } from 'powerbi-client';
import { IHttpPostMessageResponse } from 'http-post-message';
import { PowerBIEmbed } from 'powerbi-client-react';
import 'powerbi-report-authoring';

import { sampleReportUrl } from './public/constants';
import './DemoApp.css';

// Root Component to demonstrate usage of embedded component
function DemoApp (): JSX.Element {

	// PowerBI Report object (to be received via callback)
	const [report, setReport] = useState<Report>();
	const [pseudo, setPseudo] = useState<string>();

	// Track Report embedding status
	const [isEmbedded, setIsEmbedded] = useState<boolean>(false);

    // Overall status message of embedding
	const [displayMessage, setMessage] = useState(`The report is bootstrapped. Click the Embed Report button to set the access token`);

	// CSS Class to be passed to the embedded component
	const reportClass = 'report-container';

	// Pass the basic embed configurations to the embedded component to bootstrap the report on first load
    // Values for properties like embedUrl, accessToken and settings will be set on click of button
	const [sampleReportConfig, setReportConfig] = useState<models.IReportEmbedConfiguration>({
		type: 'report',
		embedUrl: undefined,
		tokenType: models.TokenType.Embed,
		accessToken: undefined,
		settings: undefined,
	});

	/**
	 * Map of event handlers to be applied to the embedded report
	 * Update event handlers for the report by redefining the map using the setEventHandlersMap function
	 * Set event handler to null if event needs to be removed
	 * More events can be provided from here
	 * https://docs.microsoft.com/en-us/javascript/api/overview/powerbi/handle-events#report-events
	 */
	const[eventHandlersMap, setEventHandlersMap] = useState<Map<string, (event?: service.ICustomEvent<any>, embeddedEntity?: Embed) => void | null>>(new Map([
		['loaded', () => console.log('Report has loaded')],
		['rendered', () => console.log('Report has rendered')],
		['error', (event?: service.ICustomEvent<any>) => {
				if (event) {


					console.error(event.detail);
				}
			},
		],
		['visualClicked', () => console.log('visual clicked')],
		['pageChanged', (event) => console.log(event)],
	]));

	useEffect(() => {
		if (report) {
			report.setComponentTitle('Embedded Report');
		}
	}, [report]);

    /**
     * Embeds report
     *
     * @returns Promise<void>
     */
	const embedReport = async (): Promise<void> => {
		console.log('Embed Report clicked');

		// Get the embed config from the service
		const reportConfigResponse = await fetch(sampleReportUrl);

		if (reportConfigResponse === null) {
			return;
		}

		if (!reportConfigResponse?.ok) {
			console.error(`Failed to fetch config for report. Status: ${ reportConfigResponse.status } ${ reportConfigResponse.statusText }`);
			return;
		}

		const reportConfig = await reportConfigResponse.json();
		console.log(reportConfig.EmbedUrl);
		// Update the reportConfig to embed the PowerBI report
		setReportConfig({
			...sampleReportConfig,
			//embedUrl: reportConfig.EmbedUrl,
			embedUrl: "https://app.powerbi.com/reportEmbed?reportId=13e17eb9-47a4-436a-9ef8-87738d04741d&config=eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9XQUJJLVVLLVNPVVRILXJlZGlyZWN0LmFuYWx5c2lzLndpbmRvd3MubmV0IiwiZW1iZWRGZWF0dXJlcyI6eyJ1c2FnZU1ldHJpY3NWTmV4dCI6dHJ1ZX19",
			//accessToken: reportConfig.EmbedToken.Token
			accessToken: "H4sIAAAAAAAEAB2TtY7FBgAE_-VaRzJTpBRmZnZnfIZnZkf591zST7Ozmr9_7Oz5Tln58-dPJR7lcbjv5zH31eAa3LXck1Ki-CK6wFto5BMoRz5fGT9r1gUKeh37HGr4SpbWdUblJ35z1PnYGllZtt8ShjublpzQk_Eaw5VkBDR2OvrE0E1hK4huA4yBcR2bpgCUcP-q5F592rJwL03CLASxWO0CofTjgq-TMWQKo-Mcv2LF7TpYGFqZDPiTRAaOYkw4fo7YM_eWGwfT0x3n7FeA4TA3XCm9hCTuAeIo9G6ydoWCXjqzWvYMdIgdMAPoVlDD2Zb3KG2r0FF9yzEeNSA1qrw6rRLuHKi2mSdHAj7fR7hIEDbfsxaiwP8c8koxzt4MBBKqRn2peDTKl1Z8u1ocCIEpTre5Zftu9vBVIQwL-SBpWtcbXyems22PXLYhQynXJZ3xbsm0q6c-B9jkB2jPgRF5IXTtIskXvi9lIPVEa8F4yjHdrW1seiNZFKLCH7kqtdxEUHfOePXvgP3iEo-9SITzDzX3koS_UinhIZskcfXZxKKchOM-r3iSvng-DD1EtuxS-tdnkvGhpO1oVNxjOLuK5tGZi7XTQni4Z7XesHLhPDbG0AcKgrcQWPBs-NBJaL3AKcuFWmj5mTgsDLEoV7LrYKwt7pNcpiOPf8itmqrYNwxV3ibF0OrmtSQYys3G9MhlmX9Qq8VvWiev764SoFea-1htrQsjyyMz1qOk6cahxboIeM728Bx5UNllj7LKb7Dq5DrE25YuyKy8WhOaho6zKSKaQj5AIxBGUp7O8rdTGXyK1lsDX6n4FnhjnQAMmcNV1EFpDtvrAGszyJudM2fSL358z8CLIcikTOj6QFylJaDDUn_9_PHDrc-8T1r1_KZTLoYdXvxBY_psYs2l1Hn3RCwqbsS4b37hAunS9yoUnoKX1OoNizFuweS5M_1g67b4idb1SyzgDru7KElpn5hndN39rOogFL2nLVttlonf1rWT0x2GIPktIUGjXGU0U-jQuDDEI5PyLa7nG76TtGvx-Qa1m4Ka57xouGkhG1Ucp3QyMKxIU_DIWSYJPB4fUllIUFXBN69PWEa5l7ixnTdjTlcXmZZ7XFDk_BtdCw2QN9OMH-jTwyuxWwXqm8ZXSJD0DC8OyWYBg6sJ896XpDc74oAnQhK6uebTmiIkpbtRJvxXUG_oDtRdn7PWL_lVUuC6Yq2TzY7XOPEPzrhgo0Y1bTh__a_5mZtqVcJfywgbgKxq4sxsZQEnp4oYyzjzP-W1nzHbj7X6xRLaqpjdAuYJJ3hsNe1JFKZQIcPLvW6l-21k56txa-NZ2V2yXaTKQF5DSm18t18m-pQUz-yr4qXsIDg4BA4aiXQ3Q_JhaniNdQX4zLN3fVDKWMW9hT17dHm0o1QN_wXT10mvlD5nK386NwIu61qGvvF0CwJfkEfL6eCyJeZEFBlfGKHKwwfKEchxPJktOccCCorDLe7OEWhJ8GLRkAEitbBmniult3DbGhO18KU3PDCEq3FKZu2tos9I-Bx0Tu9T3_bNeQNPk7RtOw2iG08Hi5F76_kCkR7YGuvMsN5vkFMt33WPLL8GamzIKN2aB-gRWzyYe1ChRw_wmff1MBuYFar_74x__gUq4qOGLgYAAA==.eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9XQUJJLVVLLVNPVVRILXJlZGlyZWN0LmFuYWx5c2lzLndpbmRvd3MubmV0IiwiZXhwIjoxNzE3NjA0Mzc3LCJhbGxvd0FjY2Vzc092ZXJQdWJsaWNJbnRlcm5ldCI6dHJ1ZX0="
		});
		setIsEmbedded(true);

		// Update the display message
		setMessage('Use the buttons above to interact with the report using Power BI Client APIs.');
	};

    /**
     * Hide Filter Pane
     *
     * @returns Promise<IHttpPostMessageResponse<void> | undefined>
     */
	const hideFilterPane = async (): Promise<IHttpPostMessageResponse<void> | undefined>  => {
		// Check if report is available or not
		if (!report) {
			setDisplayMessageAndConsole('Report not available');
			return;
		}

		// New settings to hide filter pane
		const settings = {
			panes: {
				filters: {
					expanded: false,
					visible: false,
				},
			},
		};

		try {
			const response: IHttpPostMessageResponse<void> = await report.updateSettings(settings);

			// Update display message
			setDisplayMessageAndConsole('Filter pane is hidden.');
			return response;
		} catch (error) {
			console.error(error);
			return;
		}
	};

    /**
     * Set data selected event
     *
     * @returns void
     */
	const setDataSelectedEvent = () => {
		setEventHandlersMap(new Map<string, (event?: service.ICustomEvent<any>, embeddedEntity?: Embed) => void | null> ([
			...eventHandlersMap,
			['dataSelected', (event) => {
				if(event){
					console.log(event.detail.dataPoints[0].identity[0].equals)
					setPseudo(event.detail.dataPoints[0].identity[0].equals)
				}
				console.log(event)


			}
			],
		]));

		setMessage('Data Selected event set successfully. Select data to see event in console.');
	}

    /**
     * Change visual type
     *
     * @returns Promise<void>
     */
	const changeVisualType = async (): Promise<void> => {
		// Check if report is available or not
		if (!report) {
			setDisplayMessageAndConsole('Report not available');
			return;
		}

		// Get active page of the report
		const activePage: Page | undefined = await report.getActivePage();

		if (!activePage) {
			setMessage('No Active page found');
			return;
		}

		try {
			// Change the visual type using powerbi-report-authoring
			// For more information: https://docs.microsoft.com/en-us/javascript/api/overview/powerbi/report-authoring-overview
			const visual = await activePage.getVisualByName('VisualContainer6');

			const response = await visual.changeType('lineChart');

			setDisplayMessageAndConsole(`The ${visual.type} was updated to lineChart.`);

			return response;
		}
		catch (error) {
			if (error === 'PowerBIEntityNotFound') {
				console.log('No Visual found with that name');
			} else {
				console.log(error);
			}
		}
	};

	/**
     * Set display message and log it in the console
     *
     * @returns void
     */
	const setDisplayMessageAndConsole = (message: string): void => {
		setMessage(message);
		console.log(message);
	}

	const controlButtons =
		isEmbedded ?
		<>
			<button onClick = { changeVisualType }>
				Change visual type</button>

			<button onClick = { hideFilterPane }>
				Hide filter pane</button>

			<button onClick = { setDataSelectedEvent }>
				Set event</button>

			<label className = "display-message">
				{ displayMessage }
			</label>
		</>
		:
		<>
			<label className = "display-message position">
				{ displayMessage }
			</label>

			<button onClick = { embedReport } className = "embed-report">
				Embed Report</button>
		</>;

	const header =
		<div className = "header">Power BI Embedded React Component Demo</div>;

	const reportComponent =
		<PowerBIEmbed
			embedConfig = { sampleReportConfig }
			eventHandlers = { eventHandlersMap }
			cssClassName = { reportClass }
			getEmbeddedComponent = { (embedObject: Embed) => {
				console.log(`Embedded object of type "${ embedObject.embedtype }" received`);
				setReport(embedObject as Report);
			} }
		/>;

	const footer =
		<div className = "footer">
			<p>This demo is powered by Power BI Embedded Analytics</p>
			<label className = "separator-pipe">|</label>
			<img title = "Power-BI" alt = "PowerBI_Icon" className = "footer-icon" src = "./assets/PowerBI_Icon.png" />
			<p>Explore our<a href = "https://aka.ms/pbijs/" target = "_blank" rel = "noreferrer noopener">Playground</a></p>
			<label className = "separator-pipe">|</label>
			<img title = "GitHub" alt = "GitHub_Icon" className = "footer-icon" src = "./assets/GitHub_Icon.png" />
			<p>Find the<a href = "https://github.com/microsoft/PowerBI-client-react" target = "_blank" rel = "noreferrer noopener">source code</a></p>
		</div>;

	return (
		<div className = "container">
			{ header }

			<div className = "controls">
				{ controlButtons }

				{ isEmbedded ? reportComponent : null }
			</div>
			<div>
				{pseudo ? "Selected Pseudo: " :""}
				{pseudo}
			</div>
			{ footer }
		</div>
	);
}

export default DemoApp;