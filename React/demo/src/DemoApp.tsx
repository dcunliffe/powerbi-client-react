// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import 'bootstrap/dist/css/bootstrap.min.css'
import React, { useState, useEffect } from 'react';
import { models, Report, Embed, service, Page } from 'powerbi-client';
import { IHttpPostMessageResponse } from 'http-post-message';
import { PowerBIEmbed } from 'powerbi-client-react';
import 'powerbi-report-authoring';

import { sampleReportUrl } from './public/constants';
import './DemoApp.css';
import { Card, Col, Container, Form, FormSelect, Nav, NavDropdown, Navbar, Row, Table, Toast, ToastContainer } from 'react-bootstrap';

class ReidentifcationData {
	skid!: string;
	nhs!: string;
}

// Root Component to demonstrate usage of embedded component
function DemoApp(): JSX.Element {

	// PowerBI Report object (to be received via callback)
	const [report, setReport] = useState<Report>();
	const [pseudo, setPseudo] = useState<string>();

	const [embedUrl, setEmbedUrl] = useState<string>("https://app.powerbi.com/reportEmbed?reportId=13e17eb9-47a4-436a-9ef8-87738d04741d&config=eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9XQUJJLVVLLVNPVVRILXJlZGlyZWN0LmFuYWx5c2lzLndpbmRvd3MubmV0IiwiZW1iZWRGZWF0dXJlcyI6eyJ1c2FnZU1ldHJpY3NWTmV4dCI6dHJ1ZX19");
	const [token, setEmbedToken] = useState<string>("H4sIAAAAAAAEAB2Tt660ZgAF3-W2WCInS3-x5LRkWKAjLvkjJ8vv7iv3U8wZ6fzzY6d3D9Li5--fs12eTSLMlWtNBTy4cDGIUrO4XAmtNNN-w--VXUe3mDKOTECX2NjtJgmM2iUcQtJ6CySiIMiO5E1oy_CpWVp8fI0FkfZnNd-NYdonLWnhPI0f1Fc-gblpFqEbl6TDiypJz9sxASmwn2RitW-uR4RZvrw8M3fCIFU43ozTzsY8xOYs7tq58nuq6ExI9vxF_Q4PLaRxXymaj4OEGNIxRSoH0XDaKNwE1SbFfi_TLgUBtu2KgK5yxetvggcOEDdEP7_V0CIAqDOGQpW4mJ0jsXsP-0jvXJEy3eHR1IUqZxoI0qIpDWFe09rZ3HMbTjLz7DdLE0EBVDhCwrjwIw1nP3ci59-8vV1M7PQFgTDpwXKG0h2LpGzxV2hqUcb3Qif1q7CRyGqB5IyIkgWye3Xzo910YFwdJpezMasgyqi9jHeQNPN8gSqStQ9JM01H-dz2suC3kihjEc8MdGLLNUhcZmQ7pU45Od3eQPcGau2PMk-YQbGIPfAzJpOKr1mALVHE1g7M8gDPQgQxonRTm_fmjU-mfiZ37qXeix_xvXVTn4xpuDZRx6KdtI-X9kDzxyjS9OZGceVuuPWIXEfuQBItaFY0HAJGvvOoVa77Uly7IOTdDZ73Bl9fM2Rj67GesaDq9x5rizJacW3dtCXyzpAY0u9quxt0zCm9YXA0L44BuEShO0JJ_YyvQOuakS4tt_mi5Zgl24oV0KJjwr61MiTSgwtDj3IdyZPKj4rsBCNjlGUVos9L9_nV_ZwZ5wtfZCWs85F7oUJBc1IaXzNGicO52wH1aTAxMfFVQVaJLezio5FpvhdE_Ofnrx9-uacN6OX9e51oWlxlOzNl9ATGP8nAHIvyIOOQlkVqPetEe17zcWThBtCPS-bVMz8z6aKdWlW0HJwxy5y7QolQEWOECTFGHWO0aVl5Xh5fC8PvpTbVhR9VGkGHrdroXi6OjseTELrRWfbPiwi4dVC8lvxwSwa-82CO7OqRUKazSVQJg964-4IZgLlRrdvMewQfYsfaRwLLOgQr6rFGqGf0t3PpaWL2DXvdsuyxx_iC54Byy0zKer7RFW1O8Vd6DiP6HrX4ma156SPKwPR4VuWqt7XmHWPf8GRC916iNIEf06nXoPR6vrDkPtXcrs1GGvYPX1Xki-EMjRFz3jl6XHJmCeb8bRIRUvemz-vP_5nvqS4XNfytDHE3dlhgLq_MfOD0zFb8gM__Ka_5jum2L-UvpvfYCTvw4733yVQYCSveENMP-Ks_i-y0BV_DNHDoAQGfu6DNeMwTT-EW-4P5vAtl9LzYjYJRLyFRCKPTJhuh-kiTIxmIcPy6mpV7YTuaqxyKKK0abbRwPaue1JaQtmAEXih2WovAXydUctd_mAIJuk8UsKv85d4GhN4fpSfJog95atLxN06HdUQvlMY4BOWIlyrh810DnQig6gEpsWA5dd0BXho2lAYYGk86d2y06c2cUNjW8dXOF6Wz7sgNiPDIqrROwLz4m8XSPayXiYLB1byQiT8OwY_vkj6GACJ3T1EsDnyS-raNg9-iRh1U8OaT3Gekg-QiTuXXlsNBXRUIzMPOb-Z__wPAKgI5LgYAAA==.eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9XQUJJLVVLLVNPVVRILXJlZGlyZWN0LmFuYWx5c2lzLndpbmRvd3MubmV0IiwiZXhwIjoxNzE3NjExMjg3LCJhbGxvd0FjY2Vzc092ZXJQdWJsaWNJbnRlcm5ldCI6dHJ1ZX0=");
	const [page, setPage] = useState<string>("home");
	const [reidentifcationData, setReidentifcationData] = useState<Array<ReidentifcationData>>([]);
	const [reidReason, setReidReason] = useState<string>("");
	const reasons = ["Direct Care", "Research", "Clinical", "Audit", "Service Improvement"];


	const postbody = " { \"datasets\": [ { \"id\": \"d2f76d05-5b9c-44f1-b8f3-5294636a72f9\"} ],\"reports\": [ { \"id\": \"13e17eb9-47a4-436a-9ef8-87738d04741d\" }]}";
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
	const [eventHandlersMap, setEventHandlersMap] = useState<Map<string, (event?: service.ICustomEvent<any>, embeddedEntity?: Embed) => void | null>>(new Map([
		['loaded', (event?: service.ICustomEvent<any>) => {
			console.log(event);
			console.log('Report has loaded');
		}],
		['rendered', (event?: service.ICustomEvent<any>) => {
			console.log(event);
			console.log('Report has rendered');
			setDataSelectedEvent();
		}],
		['error', (event?: service.ICustomEvent<any>) => {
			if (event) {
				console.error(event.detail);
			}
		}],
		['visualClicked', () => console.log('visual clicked')],
		['pageChanged', (event) => console.log(event)]
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
		//console.log(reportConfig.EmbedUrl);
		// Update the reportConfig to embed the PowerBI report
		setReportConfig({
			...sampleReportConfig,
			embedUrl: embedUrl,
			accessToken: token
		});
		setIsEmbedded(true);

		// Update the display message
		setMessage('Use the buttons above to interact with the report using Power BI Client APIs.');
	};

	/**
	 * Set data selected event
	 *
	 * @returns void
	 */
	const setDataSelectedEvent = () => {
		setEventHandlersMap(new Map<string, (event?: service.ICustomEvent<any>, embeddedEntity?: Embed) => void | null>([
			...eventHandlersMap,
			['dataSelected', (event) => {
				if (event && event.detail.dataPoints[0]) {
					console.log(event.detail.dataPoints[0].identity[0].equals)
					setPseudo(event.detail.dataPoints[0].identity[0].equals)

					var skid = event.detail.dataPoints[0].identity[0].equals;
					var nhs = Math.random().toString().substr(2, 8);;


					setReidentifcationData(reidentifcationData => {
						if(reidentifcationData.find(x => x.skid === skid)) {
							return reidentifcationData;
						}	

						return [...reidentifcationData, { skid: skid, nhs: nhs }]
					})
				}
				console.log(event)
			}
			],
		]));

		setMessage('Data Selected event set successfully. Select data to see event in console.');
	}

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
			</>
			:
			<>
				<Form>
					<Form.Group>
						<Form.Label>Embed Url</Form.Label>
						<Form.Control type="text" value={embedUrl} onChange={(e) => setEmbedUrl(e.target.value)} />
						<a href="https://learn.microsoft.com/en-us/rest/api/power-bi/reports/get-report#code-try-0" target="_blank">Link</a>
					</Form.Group>
					<Form.Group>
						<Form.Label>Token</Form.Label>
						<Form.Control as="textarea" rows={5} value={token} onChange={(e) => setEmbedToken(e.target.value)} />
						<a href='https://learn.microsoft.com/en-us/rest/api/power-bi/embed-token/generate-token#code-try-0' target="_blank">Link</a>
						{postbody}
					</Form.Group>
				</Form>
				<button onClick={embedReport} className="embed-report">
					Launch Report</button>
				<div>

				</div>
			</>;

	const header =
		<div className="header">ISL Re-Identification Service</div>;

	const reportComponent =
		<PowerBIEmbed
			embedConfig={sampleReportConfig}
			eventHandlers={eventHandlersMap}
			cssClassName={reportClass}
			getEmbeddedComponent={(embedObject: Embed) => {
				console.log(`Embedded object of type "${embedObject.embedtype}" received`);
				setReport(embedObject as Report);
			}}
		/>;

	const footer =
		<div className="footer">

		</div>;

	const getContent = () => {
		switch (page) {
			case "home":
				return (
					<Row>
						<Col>
							<Card onClick={() => setPage("Report")}>
								<Card.Body>Report Based ReIdentifcation</Card.Body>
							</Card>
						</Col>
						<Col>
							<Card onClick={() => setPage("CSV")}>
								<Card.Body>CSV File ReIdentifcation</Card.Body>
							</Card>
						</Col>
						<Col>
							<Card onClick={() => setPage("List")}>
								<Card.Body>List ReIdentification</Card.Body>
							</Card>
						</Col>
					</Row>
				);

			case "Report":
				return (
					<div className="container">
						{controlButtons}
						{isEmbedded ? reportComponent : null}
					</div>
				);
			case "CSV":
				return (
					<div className="container">
						CSV
					</div>
				);
			case "List":
				return (
					<div className="container">
						List
					</div>
				);
			default:
				return (
					<div className="container">
						OOPS
					</div>
				);
		}

	}

	return (
		<>
			<Navbar className="bg-body-tertiary">
				<Container>
					<Navbar.Brand href="#home">ISL Re-Identification Service</Navbar.Brand>
					<Navbar.Collapse id="basic-navbar-nav">
						<Nav>
							<NavDropdown title="Re-identification Services" id="basic-nav-dropdown">
								<Nav.Link href="#home">Reports</Nav.Link>
								<Nav.Link href="#home">CSV File</Nav.Link>
								<Nav.Link href="#home">Comma Seperated List</Nav.Link>
							</NavDropdown>
						</Nav>
					</Navbar.Collapse>
				</Container>
			</Navbar>
			<Container>
				{getContent()}
				{isEmbedded &&
					<ToastContainer position='bottom-end'>
						<Toast>
							<Toast.Header>
								<img src="holder.js/20x20?text=%20" className="rounded me-2" alt="" />
								<strong className="me-auto">Reidentifcation</strong>
							</Toast.Header>
							<Toast.Body>
								{reidReason == "" ? <>
									Select Reason for Re-Identification
									<Form>
										<Form.Group>
											<Form.Label>Reason</Form.Label>
											<FormSelect onChange={(e) => {
												setReidReason(e.target.value);
											}}
											>
												<option value={""}>---Select Reason---</option>
												{reasons.map((reason, index) => {
													return (
														<option key={index} value={reason}>{reason}</option>
													);
												})}
											</FormSelect>
										</Form.Group>
									</Form>
								</> :
									<Table>
										<thead>
											<tr>
												<th>SKID</th>
												<th>NHS Number</th>
											</tr>
										</thead>
										<tbody>
											{reidentifcationData?.map((data, index) => {
												return (
													<tr key={index}>
														<td>{data.skid}</td>
														<td>{data.nhs}</td>
													</tr>
												);
											})}
										</tbody>
									</Table>
								}
							</Toast.Body>
						</Toast>
					</ToastContainer>
				}
			</Container>
		</>
	);
}

export default DemoApp;

/*<div className = "container">

			<div className = "controls">
				{ controlButtons }

				{ isEmbedded ? reportComponent : null }
			</div>
			<div>
				{pseudo ? "Selected Pseudo: " :""}
				{pseudo}
			</div>
			{ footer }
		</div>*/