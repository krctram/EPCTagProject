import * as React from 'react';
import styles from '../Projects.module.scss';
import { IPerformanceRatingScaleProps } from './IPerformanceRatingScaleProps';
import { IPerformanceRatingScaleState } from './IPerformanceRatingScaleState';

export default class PerformanceRatingScale extends React.Component<IPerformanceRatingScaleProps, IPerformanceRatingScaleState>
{
    constructor(props: any) {
        super(props);
        this.state = {
            AppContext: props.AppContext,
            IsLoading: false
        };
    }

    public render(): React.ReactElement<IPerformanceRatingScaleProps> {
        return (
            this.props.IsLoading == true ? <React.Fragment ></React.Fragment> :
                <div className={styles.sectionContainer}>
                    <div className={styles.sectionHeader}>
                        <div className={styles.colHeader100}>
                            <span className={styles.subTitle}>Performance Rating Instructions</span>
                        </div>
                    </div>
                    <div className={styles.sectionNotes}>
                        Rate each behavioral statement using the scale provided in the drop-down field (scale definitions provided below).
                    </div>
                    <div className={styles.sectionContent}>
                        <table className={styles.ratingtable}>
                            <tbody>
                                <tr>
                                    <td className={styles.colsrno}>4</td>
                                    <td className={styles.colDetailedDescription}>Significantly exceeds expectations for level</td>
                                </tr>
                                <tr>
                                    <td className={styles.colsrno}>3</td>
                                    <td className={styles.colDetailedDescription}>Proficient for level.</td>
                                </tr>
                                <tr>
                                    <td className={styles.colsrno}>2</td>
                                    <td className={styles.colDetailedDescription}>Progressing toward expectations for level.</td>
                                </tr>
                                <tr>
                                    <td className={styles.colsrno}>1</td>
                                    <td className={styles.colDetailedDescription}>Does not meet expectations for level.</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>

        );
    }
}