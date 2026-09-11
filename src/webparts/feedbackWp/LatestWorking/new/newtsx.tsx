import * as React from 'react';
import styles from './FeedbackWp.module.scss';
import { IFeedbackWpProps } from './IFeedbackWpProps';
import { FeedbackService, IFeedbackPrompt } from '../services/FeedbackService';
import { getFeedbackStrings, IFeedbackStrings } from '../services/FeedbackStrings';

type FlowState = 'loading' | 'voting' | 'commentPending' | 'submitted' | 'inactive';
type VoteValue = 'yes' | 'no' | undefined;

interface IFeedbackWpState {
  flow: FlowState;
  prompt: IFeedbackPrompt | undefined;
  vote: VoteValue;
  comment: string;
  isSubmitting: boolean;
  error: string | undefined;
}

// Main Feedback webpart component. Shows a prompt question (from
// FeedbackPrompts list), user picks Yes/No, can add optional comment, and
// on submit we write everything to FeedbackResponses list.
//
// One thing to keep in mind - we are not persisting "already submitted"
// anywhere. State is only in memory (React state), so refreshing page
// resets everything and user can submit again. This is intentional as
// per requirement, not a bug.
export default class FeedbackWp extends React.Component<IFeedbackWpProps, IFeedbackWpState> {
  private service: FeedbackService;

  constructor(props: IFeedbackWpProps) {
    super(props);
    this.service = new FeedbackService(props.context, props.promptsListName, props.responsesListName);
    this.state = {
      flow: 'loading',
      prompt: undefined,
      vote: undefined,
      comment: '',
      isSubmitting: false,
      error: undefined
    };
  }

  public componentDidMount(): void {
    if (!this.props.isActive) {
      this.setState({ flow: 'inactive' });
      return;
    }
    this.loadPrompt().catch(() => {
      // Not doing anything here on purpose - loadPrompt() already handles
      // its own try/catch internally and sets flow to 'inactive' with
      // error message. This outer catch is only added because lint rule
      // (no-floating-promises) was complaining, so just satisfying that.
    });
  }

  private getStrings(): IFeedbackStrings {
    return getFeedbackStrings(this.props.language);
  }

  private getErrorMessage = (err: unknown): string => {
    if (err instanceof Error) {
      return err.message;
    }

    return 'Something went wrong.';
  };

  private loadPrompt = async (): Promise<void> => {
    try {
      const prompt = await this.service.getActivePrompt(this.props.language);
      if (!prompt) {
        this.setState({ flow: 'inactive' });
        return;
      }
      this.setState({ flow: 'voting', prompt });
    } catch (err) {
      // Logging to console so at least we can debug from browser dev tools
      // if something goes wrong in production.
      // eslint-disable-next-line no-console
      console.error('FeedbackWp: failed to load active prompt', err);
      this.setState({ flow: 'inactive', error: this.getErrorMessage(err) });
    }
  };

  private onVoteClick = (vote: VoteValue): void => {
    this.setState({ vote, flow: 'commentPending' });
  };

  private onCommentChange = (e: React.ChangeEvent<HTMLTextAreaElement>): void => {
    this.setState({ comment: e.target.value });
  };

  private onCancel = (): void => {
    this.setState({ vote: undefined, comment: '', flow: 'voting', error: undefined });
  };

  private onSubmit = async (): Promise<void> => {
    const { prompt, vote, comment } = this.state;
    if (!prompt || !vote) {
      return;
    }

    this.setState({ isSubmitting: true, error: undefined });

    try {
      await this.service.addResponse({
        Title: prompt.Title,
        Comments: comment,
        Like: vote === 'yes',
        PromptId: prompt.Id
      });

      this.setState({ flow: 'submitted', isSubmitting: false });
    } catch (err) {
      this.setState({ isSubmitting: false, error: this.getErrorMessage(err) });
    }
  };

  // React needs null here, not undefined, to render nothing - this is a
  // runtime thing, not just type-checking. Learned this the hard way after
  // getting "Minified React error #152" in production build. Please do not
  // change this to undefined again.
  // eslint-disable-next-line @rushstack/no-new-null
  public render(): React.ReactElement | null {
    const { flow, prompt, vote, comment, isSubmitting, error } = this.state;

    if (flow === 'inactive' || flow === 'loading' || !prompt) {
      // eslint-disable-next-line @rushstack/no-new-null
      return null;
    }

    const strings = this.getStrings();

    return (
      <div className={styles.feedbackCard}>
        {flow !== 'submitted' && (
          <div className={styles.headerRow}>
            <div className={styles.iconBadge}>
              <ClipboardIcon />
            </div>
            <div className={styles.questionText}>{prompt.Title}</div>

            <div className={styles.voteButtons}>
              <button
                type="button"
                className={vote === 'yes' ? `${styles.voteBtn} ${styles.voteBtnSelected}` : styles.voteBtn}
                onClick={() => this.onVoteClick('yes')}
                disabled={isSubmitting}
              >
                <ThumbsUpIcon selected={vote === 'yes'} /> {strings.YesButtonLabel}
              </button>
              <button
                type="button"
                className={vote === 'no' ? `${styles.voteBtn} ${styles.voteBtnSelected}` : styles.voteBtn}
                onClick={() => this.onVoteClick('no')}
                disabled={isSubmitting}
              >
                <ThumbsDownIcon selected={vote === 'no'} /> {strings.NoButtonLabel}
              </button>
            </div>
          </div>
        )}

        {flow === 'commentPending' && (
          <div className={styles.commentSection}>
            <textarea
              className={styles.commentBox}
              placeholder={strings.CommentPlaceholder}
              value={comment}
              onChange={this.onCommentChange}
              disabled={isSubmitting}
            />
            <div className={styles.actionButtons}>
              <button
                type="button"
                className={styles.submitBtn}
                onClick={this.onSubmit}
                disabled={isSubmitting}
              >
                {isSubmitting ? strings.SubmittingButtonLabel : strings.SubmitButtonLabel}
              </button>
              <button
                type="button"
                className={styles.cancelBtn}
                onClick={this.onCancel}
                disabled={isSubmitting}
              >
                {strings.CancelButtonLabel}
              </button>
            </div>
            {error && <div className={styles.errorText}>{error}</div>}
          </div>
        )}

        {flow === 'submitted' && (
          <div className={styles.thankYouSection}>
            <div className={styles.checkIcon}>
              <CheckIcon />
            </div>
            <div>
              <span className={styles.thankYouTitle}>{strings.ThankYouTitle}</span>{' '}
              <span className={styles.thankYouSubtitle}>
                {strings.ThankYouSubtitle}
              </span>
            </div>
          </div>
        )}
      </div>
    );
  }
}

// Icons below are all hand drawn SVG, not coming from any icon library.
// If design team gives actual icon files (svg/png) later instead of Figma
// spec to redraw manually, then better to switch to image based icons -
// steps for that are given below for reference.
//
// Step 1 - put image files here (make folder if not already there):
//   src/webparts/feedbackWp/assets/clipboard.svg
//   src/webparts/feedbackWp/assets/thumbs-up.svg
//   src/webparts/feedbackWp/assets/thumbs-down.svg
//   src/webparts/feedbackWp/assets/check.svg
//
// Step 2 - import at top of this file with other imports:
//   import clipboardIconUrl from './assets/clipboard.svg';
//   import thumbsUpIconUrl from './assets/thumbs-up.svg';
//   import thumbsDownIconUrl from './assets/thumbs-down.svg';
//   import checkIconUrl from './assets/check.svg';
//
// Step 3 - change each function below to use <img> tag instead, like:
//   function ClipboardIcon(): JSX.Element {
//     return <img src={clipboardIconUrl} width={20} height={20} alt="" />;
//   }
//
// One thing to note - currently Yes/No icon changes to white color when
// selected (see stroke prop using selected variable below). A plain image
// file cannot do this color change on its own. So if switching to images,
// need two separate files per icon (normal + selected state), then pick
// correct one based on selected value.

function ClipboardIcon(): JSX.Element {
  return (
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none">
      <rect x="5" y="4" width="14" height="17" rx="2" stroke="currentColor" strokeWidth="1.6" />
      <rect x="9" y="2.5" width="6" height="3" rx="1" fill="currentColor" />
      <line x1="8" y1="9" x2="16" y2="9" stroke="currentColor" strokeWidth="1.4" />
      <line x1="8" y1="12.5" x2="16" y2="12.5" stroke="currentColor" strokeWidth="1.4" />
      <line x1="8" y1="16" x2="13" y2="16" stroke="currentColor" strokeWidth="1.4" />
    </svg>
  );
}

function ThumbsUpIcon({ selected }: { selected: boolean }): JSX.Element {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
      <path
        d="M7 11v9H4a1 1 0 01-1-1v-7a1 1 0 011-1h3zm0 0l4.5-8a2 2 0 013.6 1.4L14 9h5a2 2 0 012 2l-1.6 7.2A2 2 0 0117.4 20H10a3 3 0 01-3-3v-6z"
        stroke={selected ? '#FFFFFF' : 'currentColor'}
        strokeWidth="1.6"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function ThumbsDownIcon({ selected }: { selected: boolean }): JSX.Element {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
      <path
        d="M17 13V4h3a1 1 0 011 1v7a1 1 0 01-1 1h-3zm0 0l-4.5 8a2 2 0 01-3.6-1.4L10 15H5a2 2 0 01-2-2l1.6-7.2A2 2 0 016.6 4H14a3 3 0 013 3v6z"
        stroke={selected ? '#FFFFFF' : 'currentColor'}
        strokeWidth="1.6"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function CheckIcon(): JSX.Element {
  return (
    <svg width="28" height="28" viewBox="0 0 24 24" fill="none">
      <circle cx="12" cy="12" r="10" stroke="#2E7D32" strokeWidth="1.6" />
      <path d="M8 12.5l2.5 2.5L16 9.5" stroke="#2E7D32" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}
