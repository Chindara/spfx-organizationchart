import * as React from 'react';
import { IOrganizationChartProps } from './IOrganizationChartProps';
import { IOrganizationChartState } from './IOrganizationChartState';
import {
	Icon,
	IIconStyles,
	IPersonaProps,
	IPersonaSharedProps,
	IPersonaStyles,
	Persona,
	PersonaSize,
} from 'office-ui-fabric-react';
import UserService from '../../../services/UserService';
import styles from './OrganizationChart.module.scss';

const personaStyles: Partial<IPersonaStyles> = {
	root: { margin: '0 0 10px 0' },
};
const iconStyles: Partial<IIconStyles> = { root: { marginRight: 5 } };

export default class OrganizationChart extends React.Component<IOrganizationChartProps, IOrganizationChartState> {
	private userService: UserService;
	constructor(props: IOrganizationChartProps) {
		super(props);
		this.state = {
			Me: null,
			Manager: null,
			Reports: null,
		};

		this.userService = new UserService();
	}

	public async componentDidMount(): Promise<void> {
		await this.getData();
	}

	public getData = async (): Promise<void> => {
		const meResponse: IPersonaSharedProps = await this.userService.getMe(this.props.context);
		if (meResponse) {
			this.setState({
				Me: meResponse,
			});
		}

		const managerResponse: IPersonaSharedProps = await this.userService.getManager(this.props.context);
		if (managerResponse) {
			this.setState({
				Manager: managerResponse,
			});
		}

		const reportsResponse: IPersonaSharedProps[] = await this.userService.getDirectReports(this.props.context);
		if (reportsResponse) {
			this.setState({
				Reports: reportsResponse,
			});
		}
	};

	public _onRenderSecondaryText(props: IPersonaProps): JSX.Element {
		return (
			<div>
				<Icon iconName='Suitcase' styles={iconStyles} />
				<i className={styles.secondaryText}>{props.secondaryText}</i>
			</div>
		);
	}

	public _onRenderPrimaryText(props: IPersonaProps): JSX.Element {
		return (
			<div>
				<span className={styles.primaryText}>{props.text}</span>
			</div>
		);
	}

	public render(): React.ReactElement<IOrganizationChartProps> {
		const users = this.state.Reports;

		return (
			<>
				<Persona
					{...this.state.Manager}
					size={PersonaSize.size48}
					onRenderPrimaryText={this._onRenderPrimaryText}
					onRenderSecondaryText={this._onRenderSecondaryText}
					styles={personaStyles}
				/>
				<div className={styles.me}>
					<Persona
						{...this.state.Me}
						size={PersonaSize.size48}
						onRenderPrimaryText={this._onRenderPrimaryText}
						onRenderSecondaryText={this._onRenderSecondaryText}
						styles={personaStyles}
					/>
				</div>
				{users !== null ? (
					<div>
						{users.map((user, index) => (
							<div className={styles.directReports} key={index}>
								<Persona
									{...user}
									size={PersonaSize.size48}
									onRenderPrimaryText={this._onRenderPrimaryText}
									onRenderSecondaryText={this._onRenderSecondaryText}
									styles={personaStyles}
								/>
							</div>
						))}
					</div>
				) : null}
			</>
		);
	}
}
