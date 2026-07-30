import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import * as XLSX from "xlsx";
import {
	Upload,
	Download,
	Maximize,
	Minimize,
	RotateCcw,
	Shuffle,
	ChevronLeft,
	ChevronRight,
	CheckCircle2,
	XCircle,
	BookOpen,
	FileSpreadsheet,
	Trash2,
	Save,
	ListCollapse,
	FolderPlus,
	LayoutGrid,
	Pencil,
	Search,
	Target,
	BarChart3,
	Layers3,
	Type,
	Clock3,
	Check,
	Square,
	CheckSquare,
	Play,
	Settings2,
	X,
} from "lucide-react";
import React from "react";

type ProgressInfo = {
	attempts: number;
	correct: number;
	lastReviewed: string | null;
	nextReview: string | null;
	intervalDays: number;
	ease: number;
};

type Card = {
	id: number;
	question: string;
	correctAnswers: string[];
	options: string[];
	progress: ProgressInfo;
};

type CardSet = {
	id: string;
	name: string;
	createdAt: string;
	cards: Card[];
};

type QuizSettings = {
	minAccuracy: number;
	maxAccuracy: number;
	excludePerfect: boolean;
	dueOnly: boolean;
	randomize: boolean;
};

type QuizState = {
	index: number;
	score: number;
	selected: string[];
	submitted: boolean;
};

type SavedQuizSession = {
	version: number;
	activeSetId: string;
	quizDeckIds: number[];
	index: number;
	score: number;
	selected: string[];
	submitted: boolean;
	settings: QuizSettings;
	selectedCardIds: number[];
	quizStarted: boolean;
	savedAt: string;
};

type TypingState = {
	index: number;
	value: string;
	submitted: boolean;
	correct: boolean | null;
};

const SET_STORAGE_KEY = "quizlet_style_app_sets_v3";
const QUIZ_SESSION_STORAGE_KEY = "quizlet_style_app_active_quiz_v1";
const BACKUP_APP_NAME = "quizlet-style-webapp";
const BACKUP_VERSION = 1;

function shuffleArray<T>(array: T[]): T[] {
	const shuffled = [...array];

	for (let i = shuffled.length - 1; i > 0; i--) {
		const j = Math.floor(Math.random() * (i + 1));

		[shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
	}

	return shuffled;
}

type QuizAppBackup = {
	app: string;
	version: number;
	exportedAt: string;
	data: {
		sets: string | null;
		activeQuiz: string | null;
	};
};

/**
 * Export the app's localStorage data to a JSON file.
 */
function exportQuizProgress() {
	const backup: QuizAppBackup = {
		app: BACKUP_APP_NAME,
		version: BACKUP_VERSION,
		exportedAt: new Date().toISOString(),
		data: {
			sets: localStorage.getItem(SET_STORAGE_KEY),
			activeQuiz: localStorage.getItem(QUIZ_SESSION_STORAGE_KEY),
		},
	};

	const json = JSON.stringify(backup, null, 2);

	const blob = new Blob([json], {
		type: "application/json",
	});

	const url = URL.createObjectURL(blob);

	const link = document.createElement("a");

	const date = new Date().toISOString().slice(0, 10);

	link.href = url;
	link.download = `quizlet-backup-${date}.json`;

	document.body.appendChild(link);

	link.click();

	document.body.removeChild(link);

	URL.revokeObjectURL(url);
}

/**
 * Import a previously exported backup file.
 */
function importQuizProgress(file: File): Promise<void> {
	return new Promise((resolve, reject) => {
		const reader = new FileReader();

		reader.onload = () => {
			try {
				const text = reader.result;

				if (typeof text !== "string") {
					throw new Error("Could not read the backup file.");
				}

				const backup = JSON.parse(text);

				// Validate basic backup structure
				if (!backup || typeof backup !== "object") {
					throw new Error("This is not a valid backup file.");
				}

				if (backup.app !== BACKUP_APP_NAME) {
					throw new Error("This backup was not created by this quiz app.");
				}

				if (backup.version !== BACKUP_VERSION) {
					throw new Error(
						"This backup was created by an incompatible version of the app.",
					);
				}

				if (!backup.data || typeof backup.data !== "object") {
					throw new Error("The backup file does not contain valid data.");
				}

				const { sets, activeQuiz } = backup.data;

				// Validate that sets is either null or valid JSON
				if (sets !== null) {
					if (typeof sets !== "string") {
						throw new Error("The backup contains invalid deck data.");
					}

					const parsedSets = JSON.parse(sets);

					if (!Array.isArray(parsedSets)) {
						throw new Error("The backup contains invalid deck data.");
					}
				}

				// Validate active quiz JSON
				if (activeQuiz !== null) {
					if (typeof activeQuiz !== "string") {
						throw new Error("The backup contains invalid quiz data.");
					}

					JSON.parse(activeQuiz);
				}

				// Ask user to confirm before overwriting local data
				const confirmed = window.confirm(
					"Import this quiz backup?\n\n" +
						"This will replace the quiz data currently stored on this device.",
				);

				if (!confirmed) {
					resolve();
					return;
				}

				// Restore deck data
				if (sets !== null) {
					localStorage.setItem(SET_STORAGE_KEY, sets);
				} else {
					localStorage.removeItem(SET_STORAGE_KEY);
				}

				// Restore active quiz session
				if (activeQuiz !== null) {
					localStorage.setItem(QUIZ_SESSION_STORAGE_KEY, activeQuiz);
				} else {
					localStorage.removeItem(QUIZ_SESSION_STORAGE_KEY);
				}

				resolve();
			} catch (error) {
				reject(
					error instanceof Error
						? error
						: new Error("Could not import the backup."),
				);
			}
		};

		reader.onerror = () => {
			reject(new Error("Could not read the backup file."));
		};

		reader.readAsText(file);
	});
}
const defaultQuizSettings: QuizSettings = {
	minAccuracy: 0,
	maxAccuracy: 100,
	excludePerfect: true,
	dueOnly: false,
	randomize: true,
};

const sampleSet: CardSet = {
	id: "sample",
	name: "Sample deck",
	createdAt: new Date().toISOString(),
	cards: [
		{
			id: 1,
			question: "Upload a spreadsheet to begin",
			correctAnswers: ["Your saved decks will appear here"],
			options: [
				"Your saved decks will appear here",
				"Column A = question",
				"Column B = correct answers",
				"Column C = options",
			],
			progress: {
				attempts: 0,
				correct: 0,
				lastReviewed: null,
				nextReview: null,
				intervalDays: 0,
				ease: 2.5,
			},
		},
	],
};

function classNames(...parts: Array<string | false | null | undefined>) {
	return parts.filter(Boolean).join(" ");
}

function normalizeText(value: unknown) {
	if (value === null || value === undefined) return "";
	return String(value).trim();
}

function normalizeCompare(value: string) {
	return value.trim().toLowerCase().replace(/\s+/g, " ");
}

function uniqueStrings(values: string[]) {
	return Array.from(
		new Set(values.map((value) => value.trim()).filter(Boolean)),
	);
}

function uid() {
	return `set_${Math.random().toString(36).slice(2, 10)}_${Date.now().toString(
		36,
	)}`;
}

function shuffle<T>(items: T[]) {
	const arr = [...items];

	for (let i = arr.length - 1; i > 0; i -= 1) {
		const j = Math.floor(Math.random() * (i + 1));
		[arr[i], arr[j]] = [arr[j], arr[i]];
	}

	return arr;
}

function defaultProgress(): ProgressInfo {
	return {
		attempts: 0,
		correct: 0,
		lastReviewed: null,
		nextReview: null,
		intervalDays: 0,
		ease: 2.5,
	};
}

function ensureProgress(progress?: Partial<ProgressInfo> | null): ProgressInfo {
	return {
		...defaultProgress(),
		...(progress || {}),
	};
}

function nextReviewDate(now: Date, intervalDays: number) {
	const next = new Date(now);
	next.setDate(next.getDate() + Math.max(1, Math.round(intervalDays)));
	return next.toISOString();
}

function updateSpacedRepetition(
	progress: ProgressInfo,
	isCorrect: boolean,
): ProgressInfo {
	const now = new Date();
	const attempts = progress.attempts + 1;
	const correct = progress.correct + (isCorrect ? 1 : 0);

	if (!isCorrect) {
		return {
			attempts,
			correct,
			lastReviewed: now.toISOString(),
			nextReview: nextReviewDate(now, 1),
			intervalDays: 1,
			ease: Math.max(1.3, progress.ease - 0.2),
		};
	}

	const baseInterval = progress.intervalDays <= 0 ? 1 : progress.intervalDays;

	const nextInterval =
		baseInterval === 1 ? 3 : Math.round(baseInterval * progress.ease);

	const boundedInterval = Math.max(1, nextInterval);

	return {
		attempts,
		correct,
		lastReviewed: now.toISOString(),
		nextReview: nextReviewDate(now, boundedInterval),
		intervalDays: boundedInterval,
		ease: Math.min(2.5, progress.ease + 0.1),
	};
}

function parseTypedAnswers(value: string) {
	return uniqueStrings(
		value
			.split(/[,\n;\/]+/)
			.map((part) => normalizeCompare(part))
			.filter(Boolean),
	);
}

function parseWorkbookFile(file: File): Promise<Card[]> {
	return new Promise((resolve, reject) => {
		const reader = new FileReader();

		reader.onerror = () => reject(new Error("Failed to read file."));

		reader.onload = (e) => {
			try {
				const data = e.target?.result;

				if (!data) {
					throw new Error("Empty file.");
				}

				const workbook = XLSX.read(data, {
					type: "binary",
				});

				const firstSheetName = workbook.SheetNames[0];

				if (!firstSheetName) {
					throw new Error("No sheets found in file.");
				}

				const sheet = workbook.Sheets[firstSheetName];

				const rows: any[][] = XLSX.utils.sheet_to_json(sheet, {
					header: 1,
					blankrows: false,
				});

				const cards: Card[] = [];

				let current: Card | null = null;

				for (const row of rows) {
					const question = normalizeText(row[0]);
					const answerB = normalizeText(row[1]);
					const optionC = normalizeText(row[2]);

					if (question) {
						if (current) {
							cards.push({
								...current,
								correctAnswers: uniqueStrings(current.correctAnswers),
								options: uniqueStrings(current.options),
								progress: ensureProgress(current.progress),
							});
						}

						const initialCorrectAnswers = uniqueStrings(
							[answerB].filter(Boolean),
						);

						const initialOptions = uniqueStrings(
							[answerB, optionC].filter(Boolean),
						);

						current = {
							id: cards.length + 1,
							question,
							correctAnswers: initialCorrectAnswers,
							options: initialOptions,
							progress: defaultProgress(),
						};

						continue;
					}

					if (!current) {
						continue;
					}

					if (answerB) {
						current.correctAnswers.push(answerB);
						current.options.push(answerB);
					}

					if (optionC) {
						current.options.push(optionC);
					}
				}

				if (current) {
					cards.push({
						...current,
						correctAnswers: uniqueStrings(current.correctAnswers),
						options: uniqueStrings(current.options),
						progress: ensureProgress(current.progress),
					});
				}

				const cleaned = cards
					.map((card, idx) => ({
						...card,
						id: idx + 1,
						correctAnswers: uniqueStrings(card.correctAnswers),
						options: uniqueStrings(card.options),
						progress: ensureProgress(card.progress),
					}))
					.filter(
						(card) =>
							card.question &&
							card.correctAnswers.length > 0 &&
							card.options.length > 0,
					);

				resolve(cleaned);
			} catch (error) {
				reject(
					error instanceof Error
						? error
						: new Error("Could not parse workbook."),
				);
			}
		};

		reader.readAsBinaryString(file);
	});
}

function usePersistentSets() {
	const [sets, setSets] = useState<CardSet[]>(() => {
		if (typeof window === "undefined") {
			return [sampleSet];
		}

		try {
			const raw = window.localStorage.getItem(SET_STORAGE_KEY);

			if (!raw) return [sampleSet];

			const parsed = JSON.parse(raw) as CardSet[];

			if (!Array.isArray(parsed) || !parsed.length) {
				return [sampleSet];
			}

			return parsed.map((set) => ({
				...set,
				cards: (set.cards || []).map((card) => ({
					...card,
					progress: ensureProgress(card.progress),
					correctAnswers: uniqueStrings(card.correctAnswers || []),
					options: uniqueStrings(card.options || []),
				})),
			}));
		} catch {
			return [sampleSet];
		}
	});

	useEffect(() => {
		if (typeof window === "undefined") return;

		window.localStorage.setItem(SET_STORAGE_KEY, JSON.stringify(sets));
	}, [sets]);

	return [sets, setSets] as const;
}

function FileDropzone({
	onImport,
}: {
	onImport: (name: string, cards: Card[]) => void;
}) {
	const [dragActive, setDragActive] = useState(false);

	const inputRef = useRef<HTMLInputElement | null>(null);

	const [error, setError] = useState<string | null>(null);

	const [name, setName] = useState("New deck");

	const loadFile = useCallback(
		async (file: File) => {
			setError(null);

			try {
				const cards = await parseWorkbookFile(file);

				const derivedName = file.name.replace(/\.[^/.]+$/, "") || "New deck";

				onImport(name.trim() || derivedName, cards);
			} catch (e) {
				setError(e instanceof Error ? e.message : "Could not load file.");
			}
		},
		[name, onImport],
	);

	return (
		<div
			className={classNames(
				"rounded-3xl border-2 border-dashed p-6 transition",
				dragActive
					? "border-slate-700 bg-slate-100 dark:border-slate-500 dark:bg-slate-800"
					: "border-slate-300 bg-white dark:border-slate-700 dark:bg-slate-900",
			)}
			onDragOver={(e) => {
				e.preventDefault();
				setDragActive(true);
			}}
			onDragLeave={() => setDragActive(false)}
			onDrop={(e) => {
				e.preventDefault();
				setDragActive(false);

				const file = e.dataTransfer.files?.[0];

				if (file) {
					void loadFile(file);
				}
			}}>
			<div className="flex items-start gap-4">
				<div className="rounded-2xl bg-slate-900 p-3 text-white dark:bg-white dark:text-slate-900">
					<FileSpreadsheet className="h-6 w-6" />
				</div>

				<div className="min-w-0 flex-1">
					<h2 className="text-lg font-semibold text-slate-900 dark:text-slate-100">
						Import a spreadsheet
					</h2>

					<p className="mt-1 text-sm text-slate-600 dark:text-slate-400">
						Drag and drop an .xlsx file, or choose one manually. Column A =
						question, Column B = correct answers, Column C = options.
					</p>

					<div className="mt-4 grid gap-3">
						<label className="block">
							<span className="mb-1 block text-xs font-medium uppercase tracking-[0.18em] text-slate-400">
								Deck name
							</span>

							<input
								value={name}
								onChange={(e) => setName(e.target.value)}
								className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm outline-none dark:border-slate-700 dark:bg-slate-950 dark:text-slate-100"
								placeholder="Name this set"
							/>
						</label>

						<div className="flex flex-wrap gap-3">
							<button
								type="button"
								onClick={() => inputRef.current?.click()}
								className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2 text-sm font-medium text-white dark:bg-slate-100 dark:text-slate-900">
								<Upload className="h-4 w-4" />
								Choose file
							</button>

							<input
								ref={inputRef}
								type="file"
								accept=".xlsx,.xls,.csv"
								className="hidden"
								onChange={(e) => {
									const file = e.target.files?.[0];

									if (file) {
										void loadFile(file);
									}
								}}
							/>
						</div>
					</div>

					{error ? <p className="mt-3 text-sm text-red-600">{error}</p> : null}
				</div>
			</div>
		</div>
	);
}

function ProgressPill({ label, value }: { label: string; value: string }) {
	return (
		<div className="rounded-2xl bg-slate-50 px-3 py-2 ring-1 ring-slate-200 dark:bg-slate-900 dark:ring-slate-700">
			<div className="text-[11px] uppercase tracking-[0.18em] text-slate-400">
				{label}
			</div>

			<div className="mt-1 text-sm font-semibold text-slate-900 dark:text-slate-100">
				{value}
			</div>
		</div>
	);
}
function BackupRestore() {
	const fileInputRef = React.useRef<HTMLInputElement>(null);

	const [message, setMessage] = React.useState<string | null>(null);

	const [error, setError] = React.useState<string | null>(null);

	const handleExport = () => {
		try {
			setMessage(null);
			setError(null);

			exportQuizProgress();

			setMessage("Progress exported successfully.");
		} catch (error) {
			setError(
				error instanceof Error
					? error.message
					: "Could not export your progress.",
			);
		}
	};

	const handleImportClick = () => {
		fileInputRef.current?.click();
	};

	const handleFileSelected = async (
		event: React.ChangeEvent<HTMLInputElement>,
	) => {
		const file = event.target.files?.[0];

		// Reset the input so the same file
		// can be selected again later.
		event.target.value = "";

		if (!file) {
			return;
		}

		try {
			setMessage(null);
			setError(null);

			await importQuizProgress(file);

			setMessage("Progress restored successfully. Reloading...");

			// Give the user a moment to see
			// the success message.
			setTimeout(() => {
				window.location.reload();
			}, 500);
		} catch (error) {
			setError(
				error instanceof Error
					? error.message
					: "Could not import your progress.",
			);
		}
	};

	return (
		<div className="rounded-2xl border border-slate-200 bg-white p-5 dark:border-slate-700 dark:bg-slate-900">
			<div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
				<div>
					<h2 className="font-semibold text-slate-900 dark:text-slate-100">
						Backup & Restore
					</h2>

					<p className="mt-1 text-sm text-slate-600 dark:text-slate-400">
						Export your quiz progress to a file and import it on another
						computer.
					</p>
				</div>

				<div className="flex flex-wrap gap-2">
					<button
						type="button"
						onClick={handleExport}
						className="inline-flex items-center gap-2 rounded-xl bg-slate-900 px-4 py-2 text-sm font-medium text-white hover:bg-slate-700 dark:bg-slate-100 dark:text-slate-900 dark:hover:bg-white">
						<Download className="h-4 w-4" />
						Export Progress
					</button>

					<button
						type="button"
						onClick={handleImportClick}
						className="inline-flex items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-950 dark:text-slate-200 dark:hover:bg-slate-800">
						<Upload className="h-4 w-4" />
						Import Progress
					</button>

					<input
						ref={fileInputRef}
						type="file"
						accept=".json,application/json"
						onChange={handleFileSelected}
						className="hidden"
					/>
				</div>
			</div>

			{message && (
				<p className="mt-3 text-sm text-green-600 dark:text-green-400">
					{message}
				</p>
			)}

			{error && (
				<p className="mt-3 text-sm text-red-600 dark:text-red-400">{error}</p>
			)}
		</div>
	);
}
function App() {
	const [sets, setSets] = usePersistentSets();
	const [sidebarOpen, setSidebarOpen] = useState(false);
	const quizModeRef = useRef<HTMLDivElement>(null);
	const [isQuizFullscreen, setIsQuizFullscreen] = useState(false);
	const toggleQuizFullscreen = async () => {
		if (!quizModeRef.current) {
			return;
		}

		try {
			if (!document.fullscreenElement) {
				await quizModeRef.current.requestFullscreen();
			} else {
				await document.exitFullscreen();
			}
		} catch (error) {
			console.error("Unable to toggle fullscreen:", error);
		}
	};
	const [activeSetId, setActiveSetId] = useState<string>(
		() => sets[0]?.id || sampleSet.id,
	);

	const [mode, setMode] = useState<"study" | "quiz" | "typing" | "browse">(
		"study",
	);

	const [index, setIndex] = useState(0);

	const [flipped, setFlipped] = useState(false);

	const [showOptions, setShowOptions] = useState(false);

	const [search, setSearch] = useState("");

	const [sortBy, setSortBy] = useState<"default" | "due" | "accuracy">(
		"default",
	);

	const [expandedCardId, setExpandedCardId] = useState<number | null>(null);

	const [selectedCardIds, setSelectedCardIds] = useState<number[]>([]);

	const [quizSettings, setQuizSettings] =
		useState<QuizSettings>(defaultQuizSettings);

	const [showQuizSettings, setShowQuizSettings] = useState(false);

	const [quizStarted, setQuizStarted] = useState(false);

	const [quiz, setQuiz] = useState<QuizState>({
		index: 0,
		score: 0,
		selected: [],
		submitted: false,
	});

	const [quizDeck, setQuizDeck] = useState<Card[]>([]);

	const [typing, setTyping] = useState<TypingState>({
		index: 0,
		value: "",
		submitted: false,
		correct: null,
	});

	const [renamingId, setRenamingId] = useState<string | null>(null);

	const [renameValue, setRenameValue] = useState("");

	const [resumeAvailable, setResumeAvailable] = useState(false);

	const [sessionRestored, setSessionRestored] = useState(false);

	const activeSet = useMemo(
		() => sets.find((s) => s.id === activeSetId) || sets[0] || sampleSet,
		[activeSetId, sets],
	);

	const allCards = activeSet.cards.length ? activeSet.cards : sampleSet.cards;

	const filteredCards = useMemo(() => {
		const term = normalizeCompare(search);

		let cards = allCards.filter((card) => {
			if (!term) {
				return true;
			}

			const haystack = [card.question, ...card.correctAnswers, ...card.options]
				.join(" ")
				.toLowerCase();

			return haystack.includes(term);
		});

		if (sortBy === "due") {
			cards = [...cards].sort((a, b) => {
				const aDue = a.progress.nextReview
					? new Date(a.progress.nextReview).getTime()
					: Number.POSITIVE_INFINITY;

				const bDue = b.progress.nextReview
					? new Date(b.progress.nextReview).getTime()
					: Number.POSITIVE_INFINITY;

				return aDue - bDue;
			});
		}

		if (sortBy === "accuracy") {
			cards = [...cards].sort((a, b) => {
				const aAcc = a.progress.attempts
					? a.progress.correct / a.progress.attempts
					: 0;

				const bAcc = b.progress.attempts
					? b.progress.correct / b.progress.attempts
					: 0;

				return aAcc - bAcc;
			});
		}

		return cards;
	}, [allCards, search, sortBy]);

	const cardAccuracy = useCallback(
		(card: Card) =>
			card.progress.attempts
				? Math.round((card.progress.correct / card.progress.attempts) * 100)
				: 0,
		[],
	);

	const cardIsDue = useCallback(
		(card: Card) =>
			!card.progress.nextReview ||
			new Date(card.progress.nextReview).getTime() <= Date.now(),
		[],
	);

	const quizCandidateCards = useMemo(() => {
		const selectedMode = selectedCardIds.length > 0;

		const sourceCards = selectedMode
			? allCards.filter((card) => selectedCardIds.includes(card.id))
			: allCards;

		return sourceCards.filter((card) => {
			const accuracy = cardAccuracy(card);

			if (
				quizSettings.excludePerfect &&
				card.progress.attempts > 0 &&
				accuracy === 100
			) {
				return false;
			}

			if (accuracy < quizSettings.minAccuracy) {
				return false;
			}

			if (accuracy > quizSettings.maxAccuracy) {
				return false;
			}

			if (quizSettings.dueOnly && !cardIsDue(card)) {
				return false;
			}

			return true;
		});
	}, [allCards, selectedCardIds, quizSettings, cardAccuracy, cardIsDue]);

	const currentCard = filteredCards[index] || filteredCards[0] || allCards[0];

	const currentTypingCard = allCards[typing.index] || allCards[0];

	const activeQuizCard = quizDeck[quiz.index];

	const shuffledQuizOptions = useMemo(() => {
		if (!activeQuizCard) {
			return [];
		}

		return shuffleArray(activeQuizCard.options);
	}, [activeQuizCard?.id]);

	const isMultiSelect =
		activeQuizCard?.correctAnswers.length > 1 ||
		/\(choose\s+two\)|\(choose\s+all\s+that\s+apply\)/i.test(
			activeQuizCard?.question || "",
		);

	const stats = useMemo(() => {
		const total = allCards.length;

		const totalOptions = allCards.reduce(
			(sum, card) => sum + card.options.length,
			0,
		);

		const reviewed = allCards.filter(
			(card) => card.progress.attempts > 0,
		).length;

		const mastered = allCards.filter((card) => {
			const accuracy = card.progress.attempts
				? card.progress.correct / card.progress.attempts
				: 0;

			return accuracy >= 0.8 && card.progress.attempts >= 3;
		}).length;

		const dueNow = allCards.filter((card) => cardIsDue(card)).length;

		const avgAccuracy = total
			? Math.round(
					(allCards.reduce(
						(sum, card) =>
							sum +
							(card.progress.attempts
								? card.progress.correct / card.progress.attempts
								: 0),
						0,
					) /
						total) *
						100,
				)
			: 0;

		return {
			total,
			totalOptions,
			reviewed,
			mastered,
			dueNow,
			avgAccuracy,
		};
	}, [allCards, cardIsDue]);

	const saveQuizSession = useCallback(
		(overrides?: Partial<SavedQuizSession>) => {
			if (typeof window === "undefined") {
				return;
			}

			if (!quizStarted || !quizDeck.length) {
				return;
			}

			const session: SavedQuizSession = {
				version: 1,
				activeSetId,
				quizDeckIds: quizDeck.map((card) => card.id),
				index: quiz.index,
				score: quiz.score,
				selected: quiz.selected,
				submitted: quiz.submitted,
				settings: quizSettings,
				selectedCardIds,
				quizStarted: true,
				savedAt: new Date().toISOString(),
				...overrides,
			};

			window.localStorage.setItem(
				QUIZ_SESSION_STORAGE_KEY,
				JSON.stringify(session),
			);

			setResumeAvailable(true);
		},
		[activeSetId, quiz, quizDeck, quizSettings, selectedCardIds, quizStarted],
	);

	const clearQuizSession = useCallback(() => {
		if (typeof window !== "undefined") {
			window.localStorage.removeItem(QUIZ_SESSION_STORAGE_KEY);
		}

		setResumeAvailable(false);
	}, []);

	useEffect(() => {
		const handleFullscreenChange = () => {
			setIsQuizFullscreen(document.fullscreenElement === quizModeRef.current);
		};

		document.addEventListener("fullscreenchange", handleFullscreenChange);

		return () => {
			document.removeEventListener("fullscreenchange", handleFullscreenChange);
		};
	}, []);

	useEffect(() => {
		if (typeof window === "undefined") {
			return;
		}

		const raw = window.localStorage.getItem(QUIZ_SESSION_STORAGE_KEY);

		if (!raw) {
			return;
		}

		try {
			const saved = JSON.parse(raw) as SavedQuizSession;

			if (
				!saved ||
				!saved.quizStarted ||
				!saved.activeSetId ||
				!saved.quizDeckIds?.length
			) {
				return;
			}

			const savedSet = sets.find((set) => set.id === saved.activeSetId);

			if (!savedSet) {
				return;
			}

			const restoredDeck = saved.quizDeckIds
				.map((id) => savedSet.cards.find((card) => card.id === id))
				.filter((card): card is Card => Boolean(card));

			if (!restoredDeck.length) {
				return;
			}

			setActiveSetId(saved.activeSetId);

			setSelectedCardIds(saved.selectedCardIds || []);

			setQuizSettings({
				...defaultQuizSettings,
				...(saved.settings || {}),
			});

			setQuizDeck(restoredDeck);

			setQuiz({
				index: Math.min(saved.index || 0, restoredDeck.length - 1),
				score: saved.score || 0,
				selected: saved.selected || [],
				submitted: Boolean(saved.submitted),
			});

			// Automatically restore a saved quiz after page refresh
			setQuizStarted(true);
			setMode("quiz");

			setResumeAvailable(true);
			setSessionRestored(true);
		} catch {
			window.localStorage.removeItem(QUIZ_SESSION_STORAGE_KEY);

			setResumeAvailable(false);
		}
	}, [sets]);

	useEffect(() => {
		if (!quizStarted || !quizDeck.length || sessionRestored) {
			return;
		}

		saveQuizSession();
	}, [
		quiz.index,
		quiz.score,
		quiz.selected,
		quiz.submitted,
		quizDeck,
		quizStarted,
		saveQuizSession,
		sessionRestored,
	]);

	useEffect(() => {
		if (!sets.length) {
			setSets([sampleSet]);

			setActiveSetId(sampleSet.id);
		}
	}, [sets, setSets]);

	useEffect(() => {
		if (!sessionRestored) {
			return;
		}

		setSessionRestored(false);
	}, [sessionRestored]);

	useEffect(() => {
		setIndex(0);
	}, [search, sortBy]);

	useEffect(() => {
		if (activeSet) {
			setIndex(0);
			setFlipped(false);
			setShowOptions(false);
			setExpandedCardId(null);
		}
	}, [activeSetId]);
	useEffect(() => {
		const handleQuizKeyDown = (event: KeyboardEvent) => {
			// Only respond when actively taking a quiz
			if (mode !== "quiz" || !quizStarted || !activeQuizCard) {
				return;
			}

			// Only respond to Spacebar
			if (event.code !== "Space") {
				return;
			}

			// Don't trigger when typing in an input, textarea, or select
			const target = event.target as HTMLElement;

			if (
				target.tagName === "INPUT" ||
				target.tagName === "TEXTAREA" ||
				target.tagName === "SELECT"
			) {
				return;
			}

			// Prevent the page from scrolling when Spacebar is pressed
			event.preventDefault();

			// If the question hasn't been submitted yet,
			// Spacebar submits the selected answer.
			if (!quiz.submitted) {
				// Don't submit if no answer has been selected
				if (quiz.selected.length === 0) {
					return;
				}

				submitQuiz();
				return;
			}

			// If the answer has already been submitted,
			// Spacebar moves to the next question.
			nextQuiz();
		};

		window.addEventListener("keydown", handleQuizKeyDown);

		return () => {
			window.removeEventListener("keydown", handleQuizKeyDown);
		};
	}, [mode, quizStarted, activeQuizCard, quiz.submitted, quiz.selected]);

	const saveNewSet = (name: string, importedCards: Card[]) => {
		const cleanCards = importedCards.map((card, idx) => ({
			...card,
			id: idx + 1,
			correctAnswers: uniqueStrings(card.correctAnswers),
			options: uniqueStrings(card.options),
			progress: ensureProgress(card.progress),
		}));

		const newSet: CardSet = {
			id: uid(),
			name,
			createdAt: new Date().toISOString(),
			cards: cleanCards,
		};

		setSets((prev) => [newSet, ...prev.filter((s) => s.id !== sampleSet.id)]);

		setActiveSetId(newSet.id);

		setSelectedCardIds([]);

		setQuizStarted(false);
		clearQuizSession();
	};

	const updateActiveSet = (updater: (set: CardSet) => CardSet) => {
		setSets((prev) =>
			prev.map((s) => (s.id === activeSet.id ? updater(s) : s)),
		);
	};

	const updateCardProgress = (cardId: number, isCorrect: boolean) => {
		updateActiveSet((set) => ({
			...set,
			cards: set.cards.map((card) =>
				card.id === cardId
					? {
							...card,
							progress: updateSpacedRepetition(card.progress, isCorrect),
						}
					: card,
			),
		}));
	};

	const toggleCardSelection = (cardId: number) => {
		setSelectedCardIds((prev) =>
			prev.includes(cardId)
				? prev.filter((id) => id !== cardId)
				: [...prev, cardId],
		);
	};

	const allVisibleSelected =
		filteredCards.length > 0 &&
		filteredCards.every((card) => selectedCardIds.includes(card.id));

	const selectAllCards = () => {
		setSelectedCardIds((prev) =>
			Array.from(new Set([...prev, ...filteredCards.map((card) => card.id)])),
		);
	};

	const clearAllCards = () => {
		setSelectedCardIds((prev) =>
			prev.filter((id) => !filteredCards.some((card) => card.id === id)),
		);
	};

	const startQuizWithCards = (cards: Card[]) => {
		if (!cards.length) {
			return;
		}

		const deck = quizSettings.randomize ? shuffle(cards) : [...cards];

		setQuizDeck(deck);
		setQuiz({
			index: 0,
			score: 0,
			selected: [],
			submitted: false,
		});
		setQuizStarted(true);
		setMode("quiz");
		setShowQuizSettings(false);
		setSessionRestored(false);

		setTimeout(() => {
			const session: SavedQuizSession = {
				version: 1,
				activeSetId,
				quizDeckIds: deck.map((card) => card.id),
				index: 0,
				score: 0,
				selected: [],
				submitted: false,
				settings: quizSettings,
				selectedCardIds,
				quizStarted: true,
				savedAt: new Date().toISOString(),
			};

			if (typeof window !== "undefined") {
				window.localStorage.setItem(
					QUIZ_SESSION_STORAGE_KEY,
					JSON.stringify(session),
				);
			}

			setResumeAvailable(true);
		}, 0);
	};

	const startQuiz = () => {
		startQuizWithCards(quizCandidateCards);
	};

	const startCustomQuiz = () => {
		const selectedCards = allCards.filter((card) =>
			selectedCardIds.includes(card.id),
		);

		startQuizWithCards(selectedCards);
	};
	const resumeQuiz = () => {
		if (typeof window === "undefined") {
			return;
		}

		const raw = window.localStorage.getItem(QUIZ_SESSION_STORAGE_KEY);

		if (!raw) {
			setResumeAvailable(false);
			return;
		}

		try {
			const saved = JSON.parse(raw) as SavedQuizSession;

			if (
				!saved ||
				!saved.quizStarted ||
				!saved.activeSetId ||
				!saved.quizDeckIds?.length
			) {
				setResumeAvailable(false);
				return;
			}

			const savedSet = sets.find((set) => set.id === saved.activeSetId);

			if (!savedSet) {
				setResumeAvailable(false);
				return;
			}

			const restoredDeck = saved.quizDeckIds
				.map((id) => savedSet.cards.find((card) => card.id === id))
				.filter((card): card is Card => Boolean(card));

			if (!restoredDeck.length) {
				setResumeAvailable(false);
				return;
			}

			setActiveSetId(saved.activeSetId);

			setSelectedCardIds(saved.selectedCardIds || []);

			setQuizSettings({
				...defaultQuizSettings,
				...(saved.settings || {}),
			});

			setQuizDeck(restoredDeck);

			setQuiz({
				index: Math.min(saved.index || 0, restoredDeck.length - 1),
				score: saved.score || 0,
				selected: saved.selected || [],
				submitted: Boolean(saved.submitted),
			});

			setQuizStarted(true);
			setMode("quiz");
			setResumeAvailable(true);
			setSessionRestored(true);
		} catch {
			window.localStorage.removeItem(QUIZ_SESSION_STORAGE_KEY);
			setResumeAvailable(false);
		}
	};
	const abandonQuiz = () => {
		// Make absolutely sure the latest progress is saved
		saveQuizSession();

		// Leave the quiz without deleting the saved session
		setQuizStarted(false);
		setQuizDeck([]);

		setQuiz({
			index: 0,
			score: 0,
			selected: [],
			submitted: false,
		});

		// Keep the user in Quiz mode so they can see Resume Quiz
		setMode("quiz");
	};

	const goNext = () => {
		setFlipped(false);
		setShowOptions(false);

		setIndex((prev) => (prev + 1) % Math.max(filteredCards.length, 1));
	};

	const goPrev = () => {
		setFlipped(false);
		setShowOptions(false);

		setIndex(
			(prev) =>
				(prev - 1 + Math.max(filteredCards.length, 1)) %
				Math.max(filteredCards.length, 1),
		);
	};

	const resetDeck = () => {
		setIndex(0);
		setFlipped(false);
		setShowOptions(false);
		setSelectedCardIds([]);

		setTyping({
			index: 0,
			value: "",
			submitted: false,
			correct: null,
		});

		abandonQuiz();
	};

	const toggleSelection = (choice: string) => {
		if (quiz.submitted || !activeQuizCard) {
			return;
		}

		setQuiz((prev) => {
			const alreadySelected = prev.selected.includes(choice);

			const nextSelected = alreadySelected
				? prev.selected.filter((item) => item !== choice)
				: [...prev.selected, choice];

			return {
				...prev,
				selected: nextSelected,
			};
		});
	};

	const submitQuiz = () => {
		if (!activeQuizCard || quiz.submitted) {
			return;
		}

		const selected = uniqueStrings(quiz.selected);
		const correct = uniqueStrings(activeQuizCard.correctAnswers);

		const isExactMatch =
			selected.length === correct.length &&
			selected.every((item) => correct.includes(item));

		const nextScore = quiz.score + (isExactMatch ? 1 : 0);

		// Update the quiz UI immediately.
		setQuiz((prev) => ({
			...prev,
			submitted: true,
			score: nextScore,
		}));

		// Update the card's learning progress.
		updateCardProgress(activeQuizCard.id, isExactMatch);

		// Save the current quiz session.
		saveQuizSession({
			submitted: true,
			score: nextScore,
		});
	};

	const nextQuiz = () => {
		if (quiz.index + 1 >= quizDeck.length) {
			setQuizStarted(false);
			setQuizDeck([]);

			setQuiz({
				index: 0,
				score: 0,
				selected: [],
				submitted: false,
			});

			clearQuizSession();

			return;
		}

		const nextIndex = quiz.index + 1;

		setQuiz((prev) => ({
			...prev,
			index: nextIndex,
			selected: [],
			submitted: false,
		}));

		saveQuizSession({
			index: nextIndex,
			selected: [],
			submitted: false,
		});
	};

	const submitTyping = () => {
		if (!currentTypingCard || typing.submitted) {
			return;
		}

		const typed = parseTypedAnswers(typing.value);

		const correct = uniqueStrings(currentTypingCard.correctAnswers).map(
			(item) => normalizeCompare(item),
		);

		const isCorrect =
			typed.length === correct.length &&
			typed.every((item) => correct.includes(item));

		setTyping((prev) => ({
			...prev,
			submitted: true,
			correct: isCorrect,
		}));

		updateCardProgress(currentTypingCard.id, isCorrect);
	};

	const nextTyping = () => {
		if (typing.index + 1 >= allCards.length) {
			setTyping({
				index: 0,
				value: "",
				submitted: false,
				correct: null,
			});

			return;
		}

		setTyping({
			index: typing.index + 1,
			value: "",
			submitted: false,
			correct: null,
		});
	};

	const startRename = (set: CardSet) => {
		setRenamingId(set.id);
		setRenameValue(set.name);
	};

	const confirmRename = () => {
		if (!renamingId) {
			return;
		}

		const nextName = renameValue.trim();

		if (!nextName) {
			return;
		}

		setSets((prev) =>
			prev.map((s) =>
				s.id === renamingId
					? {
							...s,
							name: nextName,
						}
					: s,
			),
		);

		setRenamingId(null);
		setRenameValue("");
	};

	const deleteSet = (id: string) => {
		setSets((prev) => {
			const next = prev.filter((s) => s.id !== id);

			if (!next.length) {
				return [sampleSet];
			}

			return next;
		});

		if (activeSetId === id) {
			const nextActive = sets.find((s) => s.id !== id)?.id || sampleSet.id;

			setActiveSetId(nextActive);

			setSelectedCardIds([]);

			abandonQuiz();
		}
	};
	const selectedSet = new Set(quiz.selected);
	useEffect(() => {
		const handleStudyKeyDown = (event: KeyboardEvent) => {
			// Only run shortcuts while in Study mode
			if (mode !== "study") {
				return;
			}

			// Don't trigger shortcuts while typing in an input,
			// textarea, or select element
			const target = event.target as HTMLElement;

			if (
				target.tagName === "INPUT" ||
				target.tagName === "TEXTAREA" ||
				target.tagName === "SELECT"
			) {
				return;
			}

			if (event.key === "ArrowLeft") {
				event.preventDefault();
				goPrev();
			}

			if (event.key === "ArrowRight") {
				event.preventDefault();
				goNext();
			}

			if (event.code === "Space") {
				event.preventDefault();
				setFlipped((v) => !v);
			}
		};

		window.addEventListener("keydown", handleStudyKeyDown);

		return () => {
			window.removeEventListener("keydown", handleStudyKeyDown);
		};
	}, [mode, goPrev, goNext]);

	return (
		<div className="min-h-screen overflow-x-hidden bg-gradient-to-b from-slate-50 via-white to-slate-100 text-slate-900 dark:from-slate-950 dark:via-slate-950 dark:to-slate-900 dark:text-slate-100">
			<div className="mx-auto flex min-h-screen w-full max-w-7xl min-w-0 flex-col px-3 py-4 sm:px-6 sm:py-6 lg:px-8">
				{/* Header intentionally kept commented out for now.
				<header className="flex min-w-0 flex-col gap-4 rounded-3xl bg-white p-5 shadow-sm ring-1 ring-slate-200 dark:bg-slate-900 dark:ring-slate-800 sm:p-6 lg:flex-row lg:items-center lg:justify-between">
					<div className="min-w-0">
						<div className="flex items-center gap-2 text-sm font-medium text-slate-500 dark:text-slate-400">
							Quizlet-style study app
						</div>

						<h1 className="mt-1 break-words text-2xl font-bold tracking-tight sm:text-3xl">
							Saved sets, custom quizzes, typing tests, and spaced repetition
						</h1>

						<p className="mt-2 max-w-3xl text-sm text-slate-600 dark:text-slate-400">
							Your decks and learning progress are saved locally. Unfinished
							quizzes are also saved automatically so you can resume them later.
						</p>
					</div>
				</header>
				*/}

				<main className="mt-6 min-w-0 flex-1">
					{/* Mobile / tablet sidebar toggle */}
					<div className="mb-4 xl:hidden">
						<button
							type="button"
							onClick={() => setSidebarOpen((open) => !open)}
							className="inline-flex w-full items-center justify-between rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-medium shadow-sm dark:border-slate-800 dark:bg-slate-900">
							<span className="flex items-center gap-2">
								<FolderPlus className="h-4 w-4" />
								{sidebarOpen ? "Hide your sets" : "Show sets"}
							</span>

							<ChevronRight
								className={classNames(
									"h-4 w-4 transition-transform",
									sidebarOpen && "rotate-90",
								)}
							/>
						</button>
					</div>

					<div className="flex min-w-0 flex-col gap-6 xl:flex-row">
						<aside
							className={classNames(
								"min-w-0 space-y-6 xl:w-[320px] xl:shrink-0",
								!sidebarOpen && "hidden xl:block",
							)}>
							<div className="rounded-3xl bg-white p-5 shadow-sm ring-1 ring-slate-200 dark:bg-slate-900 dark:ring-slate-800">
								<div className="flex items-center justify-between gap-3">
									<div className="min-w-0">
										<h2 className="text-lg font-semibold">Your sets</h2>

										<p className="text-sm text-slate-500 dark:text-slate-400">
											Stored locally in this browser.
										</p>
									</div>

									<FolderPlus className="h-5 w-5 shrink-0 text-slate-400" />
								</div>

								<div className="mt-4 space-y-3">
									{sets.map((set) => {
										const active = set.id === activeSetId;

										return (
											<div
												key={set.id}
												className={classNames(
													"rounded-2xl border p-4 transition",
													active
														? "border-slate-900 bg-slate-50 dark:border-slate-100 dark:bg-slate-950"
														: "border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900",
												)}>
												<div className="flex min-w-0 items-start justify-between gap-3">
													<button
														type="button"
														onClick={() => setActiveSetId(set.id)}
														className="min-w-0 flex-1 text-left">
														<div className="flex min-w-0 items-center gap-2">
															<LayoutGrid className="h-4 w-4 shrink-0 text-slate-400" />

															{renamingId === set.id ? (
																<input
																	autoFocus
																	value={renameValue}
																	onChange={(e) =>
																		setRenameValue(e.target.value)
																	}
																	onKeyDown={(e) => {
																		if (e.key === "Enter") {
																			confirmRename();
																		}

																		if (e.key === "Escape") {
																			setRenamingId(null);
																			setRenameValue("");
																		}
																	}}
																	onClick={(e) => e.stopPropagation()}
																	className="w-full min-w-0 rounded-xl border border-slate-200 bg-white px-3 py-1.5 text-sm outline-none dark:border-slate-700 dark:bg-slate-950"
																/>
															) : (
																<div className="truncate text-sm font-semibold text-slate-900 dark:text-slate-100">
																	{set.name}
																</div>
															)}
														</div>

														<div className="mt-1 text-xs text-slate-500 dark:text-slate-400">
															{set.cards.length} cards
															{" • "}
															{new Date(set.createdAt).toLocaleDateString(
																"en-GB",
															)}
														</div>
													</button>

													<div className="flex shrink-0 items-center gap-1">
														{renamingId === set.id ? (
															<button
																type="button"
																onClick={confirmRename}
																className="rounded-lg p-2 text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
																aria-label="Save name">
																<Save className="h-4 w-4" />
															</button>
														) : (
															<button
																type="button"
																onClick={() => startRename(set)}
																className="rounded-lg p-2 text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
																aria-label="Rename set">
																<Pencil className="h-4 w-4" />
															</button>
														)}

														<button
															type="button"
															onClick={() => deleteSet(set.id)}
															className="rounded-lg p-2 text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
															aria-label="Delete set">
															<Trash2 className="h-4 w-4" />
														</button>
													</div>
												</div>
											</div>
										);
									})}
								</div>
							</div>
						</aside>
						<section className="min-w-0 overflow-hidden rounded-3xl bg-white p-4 shadow-sm ring-1 ring-slate-200 dark:bg-slate-900 dark:ring-slate-800 sm:p-6">
							<div className="flex min-w-0 flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
								<div className="min-w-0">
									<div className="text-sm text-slate-500 dark:text-slate-400">
										Active set
									</div>

									<h2 className="break-words text-2xl font-bold">
										{activeSet.name}
									</h2>

									<p className="mt-1 text-sm text-slate-600 dark:text-slate-400">
										{allCards.length} cards, {stats.totalOptions} total options.
									</p>
								</div>

								<div className="flex shrink-0 flex-wrap gap-2">
									<button
										type="button"
										onClick={resetDeck}
										className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 px-4 py-2 text-sm font-medium hover:bg-slate-50 dark:border-slate-700 dark:hover:bg-slate-800">
										<RotateCcw className="h-4 w-4" />
										Reset
									</button>
								</div>
							</div>

							<div className="mt-5 flex min-w-0 flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
								<div className="grid min-w-0 gap-2 rounded-2xl bg-slate-100 p-2 sm:grid-cols-4 dark:bg-slate-950">
									{[
										["study", "Study", BookOpen],
										["quiz", "Quiz", Shuffle],
										["typing", "Typing", Type],
										["browse", "All cards", Layers3],
									].map(([key, label, Icon]) => (
										<button
											key={String(key)}
											className={classNames(
												"rounded-2xl px-4 py-2 text-sm font-medium transition",
												mode === key
													? "bg-white shadow-sm dark:bg-slate-900"
													: "text-slate-500 dark:text-slate-400",
											)}
											onClick={() => setMode(key as typeof mode)}>
											<span className="inline-flex items-center gap-2">
												<Icon className="h-4 w-4" />
												{String(label)}
											</span>
										</button>
									))}
								</div>

								<div className="flex min-w-0 flex-col gap-2 sm:flex-row sm:items-center">
									<label className="relative min-w-0">
										<Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />

										<input
											value={search}
											onChange={(e) => setSearch(e.target.value)}
											placeholder="Search cards"
											className="w-full min-w-0 rounded-2xl border border-slate-200 bg-white py-2 pl-9 pr-4 text-sm outline-none dark:border-slate-700 dark:bg-slate-950"
										/>
									</label>

									<select
										value={sortBy}
										onChange={(e) => setSortBy(e.target.value as typeof sortBy)}
										className="rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm outline-none dark:border-slate-700 dark:bg-slate-950">
										<option value="default">Default</option>
										<option value="due">Sort by due date</option>
										<option value="accuracy">Sort by accuracy</option>
									</select>
								</div>
							</div>
							{mode === "study" ? (
								<div className="mt-6 flex min-w-0 flex-col">
									<div className="mb-4 flex flex-wrap items-center justify-between gap-3 text-sm text-slate-500 dark:text-slate-400">
										<div>
											Card {filteredCards.length ? index + 1 : 1} of{" "}
											{filteredCards.length || 1}
										</div>

										<button
											type="button"
											onClick={() => setShowOptions((v) => !v)}
											className="inline-flex items-center gap-2 rounded-xl border border-slate-200 px-3 py-1.5 text-xs font-medium dark:border-slate-700">
											<ListCollapse className="h-4 w-4" />
											{showOptions ? "Hide options" : "View all options"}
										</button>
									</div>

									<AnimatePresence mode="wait">
										<motion.button
											key={`${currentCard?.question}-${flipped}-${showOptions}`}
											initial={{
												opacity: 0,
												y: 12,
											}}
											animate={{
												opacity: 1,
												y: 0,
											}}
											exit={{
												opacity: 0,
												y: -12,
											}}
											transition={{
												duration: 0.22,
											}}
											onClick={() => setFlipped((v) => !v)}
											className="flex min-h-[280px] w-full min-w-0 flex-col items-stretch justify-center overflow-hidden rounded-3xl border border-slate-200 bg-gradient-to-br from-slate-50 to-white p-5 text-left shadow-sm dark:border-slate-800 dark:from-slate-950 dark:to-slate-900 sm:min-h-[320px] sm:p-6">
											<div className="mb-3 text-xs font-semibold uppercase tracking-[0.24em] text-slate-400">
												{flipped ? "Answer" : "Question"}
											</div>

											<div className="break-words text-xl font-semibold leading-tight text-slate-900 dark:text-slate-100 sm:text-3xl">
												{flipped
													? currentCard?.correctAnswers.join(", ")
													: currentCard?.question}
											</div>

											{showOptions ? (
												<div className="mt-6 min-w-0">
													<div className="rounded-2xl border border-slate-200 bg-white p-4 dark:border-slate-700 dark:bg-slate-950">
														<div className="mb-2 text-xs font-semibold uppercase tracking-[0.2em] text-slate-400">
															All options
														</div>

														<div className="grid min-w-0 gap-2">
															{currentCard?.options.map((option) => {
																const isCorrect =
																	currentCard.correctAnswers.includes(option);

																return (
																	<div
																		key={option}
																		className={classNames(
																			"break-words rounded-xl border px-3 py-2 text-sm",
																			isCorrect
																				? "border-green-300 bg-green-50 dark:border-green-700 dark:bg-green-950"
																				: "border-slate-200 bg-slate-50 dark:border-slate-800 dark:bg-slate-900",
																		)}>
																		{option}
																	</div>
																);
															})}
														</div>
													</div>
												</div>
											) : null}
										</motion.button>
									</AnimatePresence>

									<div className="mt-5 flex flex-wrap items-center justify-between gap-3">
										<button
											type="button"
											onClick={goPrev}
											className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 px-4 py-2 text-sm font-medium dark:border-slate-700">
											<ChevronLeft className="h-4 w-4" />
											Prev
										</button>

										<button
											type="button"
											onClick={() => setFlipped((v) => !v)}
											className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-medium text-white dark:bg-slate-100 dark:text-slate-900">
											Flip
										</button>

										<button
											type="button"
											onClick={goNext}
											className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 px-4 py-2 text-sm font-medium dark:border-slate-700">
											Next
											<ChevronRight className="h-4 w-4" />
										</button>
									</div>
								</div>
							) : mode === "quiz" ? (
								<div
									ref={quizModeRef}
									className={classNames(
										"mt-6 min-w-0 space-y-4",
										isQuizFullscreen &&
											"m-0 min-h-screen overflow-y-auto bg-white px-4 py-8 dark:bg-slate-950 sm:px-8 sm:py-12 lg:px-38 lg:py-42",
									)}>
									<div
										className={classNames(
											isQuizFullscreen && "mx-auto w-full max-w-5xl",
										)}>
										<div className="flex flex-wrap items-center gap-2">
											<button
												type="button"
												onClick={() => setShowQuizSettings((v) => !v)}
												className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 px-4 py-2 text-sm font-medium dark:border-slate-700">
												<Settings2 className="h-4 w-4" />
												Quiz settings
											</button>

											{quizStarted ? (
												<button
													type="button"
													onClick={abandonQuiz}
													className="inline-flex items-center gap-2 rounded-2xl border border-red-200 px-4 py-2 text-sm font-medium text-red-600 dark:border-red-900">
													<X className="h-4 w-4" />
													Exit quiz
												</button>
											) : null}

											<button
												type="button"
												onClick={toggleQuizFullscreen}
												className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 px-4 py-2 text-sm font-medium dark:border-slate-700">
												{isQuizFullscreen ? (
													<Minimize className="h-4 w-4" />
												) : (
													<Maximize className="h-4 w-4" />
												)}

												{isQuizFullscreen ? "Exit fullscreen" : "Fullscreen"}
											</button>
											<div className="text-sm text-slate-500 dark:text-slate-400">
												Question {quiz.index + 1} of {quizDeck.length} • Score:{" "}
												{quiz.score}
											</div>
										</div>
									</div>

									<AnimatePresence>
										{showQuizSettings ? (
											<motion.div
												initial={{
													opacity: 0,
													height: 0,
												}}
												animate={{
													opacity: 1,
													height: "auto",
												}}
												exit={{
													opacity: 0,
													height: 0,
												}}
												className="overflow-hidden">
												<div className="rounded-3xl border border-slate-200 bg-slate-50 p-4 dark:border-slate-800 dark:bg-slate-950 sm:p-6">
													<div className="flex items-center justify-between">
														<div>
															<h3 className="font-semibold">Quiz settings</h3>

															<p className="text-sm text-slate-500 dark:text-slate-400">
																Choose which cards are included.
															</p>
														</div>
													</div>

													<div className="mt-5 grid gap-5 lg:grid-cols-2">
														<label className="block">
															<span className="mb-2 block text-sm font-medium">
																Minimum accuracy: {quizSettings.minAccuracy}%
															</span>

															<input
																type="range"
																min="0"
																max="100"
																step="5"
																value={quizSettings.minAccuracy}
																onChange={(e) =>
																	setQuizSettings((prev) => ({
																		...prev,
																		minAccuracy: Number(e.target.value),
																	}))
																}
																className="w-full"
															/>
														</label>

														<label className="block">
															<span className="mb-2 block text-sm font-medium">
																Maximum accuracy: {quizSettings.maxAccuracy}%
															</span>

															<input
																type="range"
																min="0"
																max="100"
																step="5"
																value={quizSettings.maxAccuracy}
																onChange={(e) =>
																	setQuizSettings((prev) => ({
																		...prev,
																		maxAccuracy: Number(e.target.value),
																	}))
																}
																className="w-full"
															/>
														</label>
													</div>

													<div className="mt-5 grid gap-4">
														<label className="flex items-start gap-3">
															<input
																type="checkbox"
																checked={quizSettings.excludePerfect}
																onChange={(e) =>
																	setQuizSettings((prev) => ({
																		...prev,
																		excludePerfect: e.target.checked,
																	}))
																}
																className="mt-1"
															/>

															<span>
																<span className="block text-sm font-medium">
																	Exclude 100% accuracy cards
																</span>

																<span className="block text-xs text-slate-500 dark:text-slate-400">
																	Remove cards you have answered correctly every
																	time.
																</span>
															</span>
														</label>

														<label className="flex items-start gap-3">
															<input
																type="checkbox"
																checked={quizSettings.dueOnly}
																onChange={(e) =>
																	setQuizSettings((prev) => ({
																		...prev,
																		dueOnly: e.target.checked,
																	}))
																}
																className="mt-1"
															/>

															<span>
																<span className="block text-sm font-medium">
																	Due cards only
																</span>

																<span className="block text-xs text-slate-500 dark:text-slate-400">
																	Only include cards currently due for review.
																</span>
															</span>
														</label>

														<label className="flex items-start gap-3">
															<input
																type="checkbox"
																checked={quizSettings.randomize}
																onChange={(e) =>
																	setQuizSettings((prev) => ({
																		...prev,
																		randomize: e.target.checked,
																	}))
																}
																className="mt-1"
															/>

															<span>
																<span className="block text-sm font-medium">
																	Randomise quiz
																</span>

																<span className="block text-xs text-slate-500 dark:text-slate-400">
																	Shuffle cards when starting a quiz.
																</span>
															</span>
														</label>
													</div>

													<div className="mt-5 flex flex-wrap items-center gap-3">
														<div className="text-sm text-slate-500 dark:text-slate-400">
															{quizCandidateCards.length} cards match these
															settings.
														</div>

														<button
															type="button"
															onClick={startQuiz}
															disabled={!quizCandidateCards.length}
															className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-40 dark:bg-slate-100 dark:text-slate-900">
															<Play className="h-4 w-4" />
															Start quiz
														</button>
													</div>
												</div>
											</motion.div>
										) : null}
									</AnimatePresence>

									{!quizStarted ? (
										<div className="rounded-3xl border border-slate-200 bg-slate-50 p-6 dark:border-slate-800 dark:bg-slate-950">
											<div className="text-center">
												<Shuffle className="mx-auto h-10 w-10 text-slate-400" />

												<h2 className="mt-4 text-xl font-semibold">
													Ready to practise?
												</h2>

												<p className="mx-auto mt-2 max-w-xl text-sm text-slate-500 dark:text-slate-400">
													Configure your quiz using the settings above, or
													select individual cards in All cards to build a custom
													quiz.
												</p>

												<div className="mt-5 flex flex-wrap justify-center gap-3">
													{resumeAvailable ? (
														<button
															type="button"
															onClick={resumeQuiz}
															className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-5 py-3 text-sm font-medium shadow-sm hover:bg-slate-50 dark:border-slate-700 dark:bg-slate-900 dark:hover:bg-slate-800">
															<Play className="h-4 w-4" />
															Resume quiz
														</button>
													) : null}

													<button
														type="button"
														onClick={startQuiz}
														disabled={!quizCandidateCards.length}
														className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-5 py-3 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-40 dark:bg-slate-100 dark:text-slate-900">
														<Play className="h-4 w-4" />
														Start new quiz
													</button>

													<button
														type="button"
														onClick={() => setMode("browse")}
														className="rounded-2xl border border-slate-200 px-5 py-3 text-sm font-medium dark:border-slate-700">
														Select cards
													</button>
												</div>
											</div>
										</div>
									) : activeQuizCard ? (
										<div className="min-w-0 rounded-3xl border border-slate-200 bg-slate-50 p-4 dark:border-slate-800 dark:bg-slate-950 sm:p-6">
											<div className="text-xs font-semibold uppercase tracking-[0.24em] text-slate-400">
												Quiz question
											</div>

											<h2
												className="mt-3 break-words font-semibold leading-tight text-slate-900 dark:text-slate-100"
												style={{
													fontSize: isQuizFullscreen
														? "clamp(1.5rem, 5vw, 3rem)"
														: "1.5rem",
												}}>
												{activeQuizCard.question}
											</h2>

											<div className="mt-3 text-sm text-slate-600 dark:text-slate-400">
												{isMultiSelect
													? "Select all correct answers."
													: "Select the correct answer."}
											</div>

											<div className="mt-5 grid min-w-0 gap-3">
												{shuffledQuizOptions.map((choice) => {
													const isCorrect =
														activeQuizCard.correctAnswers.includes(choice);

													const isSelected = selectedSet.has(choice);

													const showResult = quiz.submitted;

													return (
														<button
															key={choice}
															type="button"
															onClick={() => toggleSelection(choice)}
															style={{
																fontSize: isQuizFullscreen
																	? "clamp(1rem, 2vw, 2.5rem)"
																	: "0.875rem",
															}}
															className={classNames(
																"flex min-w-0 items-center justify-between gap-3 rounded-2xl border px-4 py-3 text-left font-medium transition",
																isQuizFullscreen
																	? "text-lg sm:text-xl lg:text-2xl"
																	: "text-sm",
																!showResult &&
																	isSelected &&
																	"border-slate-900 bg-slate-100 dark:border-slate-100 dark:bg-slate-800",
																!showResult &&
																	!isSelected &&
																	"border-slate-200 bg-white hover:bg-slate-500 dark:border-slate-700 dark:bg-slate-900",
																showResult &&
																	isCorrect &&
																	"border-green-300 bg-green-50 dark:border-green-700 dark:bg-green-950",
																showResult &&
																	isSelected &&
																	!isCorrect &&
																	"border-red-300 bg-red-50 dark:border-red-700 dark:bg-red-950",
															)}>
															<span className="flex min-w-0 items-center gap-3">
																<span
																	className={classNames(
																		"inline-flex h-4 w-4 shrink-0 items-center justify-center rounded border",
																		isSelected
																			? "border-slate-900 bg-slate-900 text-white dark:border-slate-100 dark:bg-slate-100 dark:text-slate-900"
																			: "border-slate-300 bg-white dark:border-slate-600 dark:bg-slate-950",
																	)}>
																	{isSelected ? (
																		<Check className="h-3 w-3" />
																	) : null}
																</span>

																<span className="wrap-break-word">
																	{choice}
																</span>
															</span>

															<span className="ml-3 shrink-0">
																{showResult && isCorrect ? (
																	<CheckCircle2 className="h-5 w-5 text-green-600" />
																) : showResult && isSelected && !isCorrect ? (
																	<XCircle className="h-5 w-5 text-red-600" />
																) : null}
															</span>
														</button>
													);
												})}
											</div>

											{quiz.submitted ? (
												<div className="mt-4 rounded-2xl border border-slate-200 bg-white p-4 text-sm dark:border-slate-700 dark:bg-slate-900">
													{uniqueStrings(quiz.selected).length ===
														uniqueStrings(activeQuizCard.correctAnswers)
															.length &&
													uniqueStrings(quiz.selected).every((item) =>
														activeQuizCard.correctAnswers.includes(item),
													)
														? "Correct."
														: "Not quite. The correct answers are highlighted in green."}
												</div>
											) : null}

											<div className="mt-5 flex flex-wrap justify-end gap-3">
												{!quiz.submitted ? (
													<button
														type="button"
														onClick={submitQuiz}
														disabled={quiz.selected.length === 0}
														className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-medium text-white disabled:opacity-40 dark:bg-slate-100 dark:text-slate-900">
														Submit
														<kbd className="ml-2 rounded border border-slate-500 px-1.5 py-0.5 text-xs">
															Space
														</kbd>
													</button>
												) : (
													<button
														type="button"
														onClick={nextQuiz}
														className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-medium text-white dark:bg-slate-100 dark:text-slate-900">
														{quiz.index + 1 >= quizDeck.length
															? "Finish"
															: "Next question"}

														<kbd className="ml-2 rounded border border-slate-500 px-1.5 py-0.5 text-xs">
															Space
														</kbd>
													</button>
												)}
											</div>
										</div>
									) : null}
								</div>
							) : mode === "typing" ? (
								<div className="mt-6 min-w-0">
									<div className="mb-4 flex flex-wrap items-center justify-between gap-3 text-sm text-slate-500 dark:text-slate-400">
										<div>
											Card {allCards.length ? typing.index + 1 : 1} of{" "}
											{allCards.length || 1}
										</div>

										<div>
											{currentTypingCard?.correctAnswers.length > 1
												? "Type answers separated by commas"
												: "Type the answer"}
										</div>
									</div>

									<div className="min-w-0 rounded-3xl border border-slate-200 bg-slate-50 p-5 dark:border-slate-800 dark:bg-slate-950 sm:p-6">
										<div className="text-xs font-semibold uppercase tracking-[0.24em] text-slate-400">
											Typing test
										</div>

										<h2 className="mt-3 break-words text-xl font-semibold leading-tight sm:text-2xl">
											{currentTypingCard?.question}
										</h2>

										<textarea
											value={typing.value}
											onChange={(e) =>
												setTyping((prev) => ({
													...prev,
													value: e.target.value,
												}))
											}
											rows={4}
											placeholder="Type your answer here"
											className="mt-4 w-full min-w-0 rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none dark:border-slate-700 dark:bg-slate-900 dark:text-slate-100"
										/>

										<div className="mt-4 flex flex-wrap gap-3">
											<button
												type="button"
												onClick={submitTyping}
												className="rounded-2xl bg-slate-900 px-4 py-2 text-sm font-medium text-white dark:bg-slate-100 dark:text-slate-900">
												Check answer
											</button>

											<button
												type="button"
												onClick={nextTyping}
												className="rounded-2xl border border-slate-200 px-4 py-2 text-sm font-medium dark:border-slate-700">
												Next card
											</button>
										</div>

										{typing.submitted ? (
											<div
												className={classNames(
													"mt-4 rounded-2xl border p-4 text-sm",
													typing.correct
														? "border-green-300 bg-green-50 dark:border-green-700 dark:bg-green-950"
														: "border-red-300 bg-red-50 dark:border-red-700 dark:bg-red-950",
												)}>
												{typing.correct
													? "Correct."
													: `Incorrect. Correct answer${currentTypingCard?.correctAnswers.length > 1 ? "s are" : " is"}: ${currentTypingCard?.correctAnswers.join(", ")}`}
											</div>
										) : null}
									</div>
								</div>
							) : (
								<div className="mt-6 min-w-0 space-y-4">
									<div className="flex flex-col gap-3 rounded-3xl border border-slate-200 bg-slate-50 p-4 dark:border-slate-800 dark:bg-slate-950 sm:flex-row sm:items-center sm:justify-between">
										<div className="min-w-0">
											<div className="text-sm font-medium">
												{selectedCardIds.length} selected
											</div>

											<div className="text-sm text-slate-500 dark:text-slate-400">
												Showing {filteredCards.length} of {allCards.length}{" "}
												cards.
											</div>
										</div>

										<div className="flex flex-wrap gap-2">
											<button
												type="button"
												onClick={
													allVisibleSelected ? clearAllCards : selectAllCards
												}
												className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-2 text-sm font-medium dark:border-slate-700 dark:bg-slate-900">
												{allVisibleSelected ? (
													<Square className="h-4 w-4" />
												) : (
													<CheckSquare className="h-4 w-4" />
												)}

												{allVisibleSelected ? "Clear all" : "Select all"}
											</button>

											<button
												type="button"
												onClick={startCustomQuiz}
												disabled={!selectedCardIds.length}
												className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-40 dark:bg-slate-100 dark:text-slate-900">
												<Play className="h-4 w-4" />
												Custom quiz
											</button>
										</div>
									</div>

									<div className="space-y-3">
										{filteredCards.map((card) => {
											const selected = selectedCardIds.includes(card.id);

											const expanded = expandedCardId === card.id;

											return (
												<div
													key={card.id}
													className={classNames(
														"min-w-0 rounded-3xl border bg-white p-4 dark:bg-slate-900",
														selected
															? "border-slate-900 ring-1 ring-slate-900 dark:border-slate-100 dark:ring-slate-100"
															: "border-slate-200 dark:border-slate-800",
													)}>
													<div className="flex min-w-0 items-start gap-3">
														<button
															type="button"
															onClick={() => toggleCardSelection(card.id)}
															className="mt-1 shrink-0"
															aria-label={
																selected
																	? "Remove card from selection"
																	: "Select card"
															}>
															{selected ? (
																<CheckSquare className="h-5 w-5" />
															) : (
																<Square className="h-5 w-5 text-slate-400" />
															)}
														</button>

														<button
															type="button"
															onClick={() =>
																setExpandedCardId((prev) =>
																	prev === card.id ? null : card.id,
																)
															}
															className="min-w-0 flex-1 text-left">
															<div className="break-words text-base font-semibold">
																{card.question}
															</div>

															<div className="mt-2 text-sm text-slate-500 dark:text-slate-400">
																Accuracy {cardAccuracy(card)}% •{" "}
																{card.progress.attempts} attempts
															</div>
														</button>

														<ChevronRight
															className={classNames(
																"h-5 w-5 shrink-0 text-slate-400 transition",
																expanded && "rotate-90",
															)}
														/>
													</div>

													{expanded ? (
														<div className="mt-4 space-y-4 rounded-2xl border border-slate-200 bg-slate-50 p-4 dark:border-slate-800 dark:bg-slate-950">
															<div>
																<div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-400">
																	Correct answer(s)
																</div>

																<div className="mt-2 flex flex-wrap gap-2">
																	{card.correctAnswers.map((answer) => (
																		<span
																			key={answer}
																			className="rounded-full border border-green-300 bg-green-50 px-3 py-1 text-sm dark:border-green-700 dark:bg-green-950">
																			{answer}
																		</span>
																	))}
																</div>
															</div>

															<div>
																<div className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-400">
																	All options
																</div>

																<div className="mt-2 grid min-w-0 gap-2 sm:grid-cols-2">
																	{card.options.map((option) => (
																		<div
																			key={option}
																			className={classNames(
																				"break-words rounded-xl border px-3 py-2 text-sm",
																				card.correctAnswers.includes(option)
																					? "border-green-300 bg-green-50 dark:border-green-700 dark:bg-green-950"
																					: "border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900",
																			)}>
																			{option}
																		</div>
																	))}
																</div>
															</div>

															<div className="grid gap-2 sm:grid-cols-3">
																<div className="flex items-center gap-2 rounded-2xl bg-white px-3 py-2 text-sm dark:bg-slate-900">
																	<Target className="h-4 w-4 text-slate-400" />
																	{card.progress.correct}/
																	{card.progress.attempts} correct
																</div>

																<div className="flex items-center gap-2 rounded-2xl bg-white px-3 py-2 text-sm dark:bg-slate-900">
																	<Clock3 className="h-4 w-4 text-slate-400" />
																	{card.progress.nextReview
																		? new Date(
																				card.progress.nextReview,
																			).toLocaleString("en-GB")
																		: "Not scheduled"}
																</div>

																<div className="flex items-center gap-2 rounded-2xl bg-white px-3 py-2 text-sm dark:bg-slate-900">
																	<BarChart3 className="h-4 w-4 text-slate-400" />
																	Ease {card.progress.ease.toFixed(1)}
																</div>
															</div>
														</div>
													) : null}
												</div>
											);
										})}
									</div>
								</div>
							)}
						</section>
					</div>
					<div className="mt-6 grid min-w-0 gap-3 sm:grid-cols-2 xl:grid-cols-6">
						<ProgressPill label="Sets" value={String(sets.length)} />

						<ProgressPill label="Cards" value={String(stats.total)} />

						<ProgressPill label="Due now" value={String(stats.dueNow)} />

						<ProgressPill label="Reviewed" value={String(stats.reviewed)} />

						<ProgressPill label="Mastered" value={String(stats.mastered)} />

						<ProgressPill
							label="Avg accuracy"
							value={`${stats.avgAccuracy}%`}
						/>
					</div>
					<div className="mt-6 grid min-w-0 gap-3 sm:grid-cols-2 xl:grid-cols-2">
						<FileDropzone onImport={saveNewSet} />
						<BackupRestore />
					</div>
				</main>
			</div>
		</div>
	);
}

export default App;
