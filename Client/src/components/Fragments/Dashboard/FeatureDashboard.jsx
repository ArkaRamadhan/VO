import React from "react";

const Section = ({ span, cols, title, borderColor, children, height }) => (
  <div
    className={`${span} w-full ${height ? `h-${height}` : 'h-full'} bg-gray-50 border-b-4 border-${borderColor}-500 rounded-lg shadow p-4 flex flex-col`}
  >
    <div className="mb-3">
      <div className="flex items-center">
        <div className="flex justify-center items-center">
          <h5 className="text-xl font-bold leading-none text-gray-900 pe-1">
            {title}
          </h5>
        </div>
      </div>
    </div>
    <div className="bg-white shadow p-3 rounded-lg flex flex-col h-full">
      <div className={`grid ${cols} gap-3 mb-2 text-center h-full`}>{children}</div>
    </div>
  </div>
);

const DataItem = ({ count, label, color, className }) => (
  <dl
    className={`bg-${color}-50 p-2 rounded-lg flex flex-col items-center justify-center h-full ${className}`}
  >
    <dt
      className={`size-8 rounded-full bg-${color}-100 text-${color}-600 text-sm font-medium flex items-center justify-center mb-1`}
    >
      {count}
    </dt>
    <dd className={`text-${color}-600 text-sm font-medium`}>{label}</dd>
  </dl>
);

export const FeatureDashboard = (props) => {
  const {
    memo = [],
    beritaAcara = [],
    surat = [],
    sk = [],
    perdin = [],
    project = [],
    timelineProject = [],
    timelineDesktop = [],
    meeting = [],
    meetingList = [],
    booking = [],
    rapat = [],
    cuti = [],
    masuk = [],
    keluar = [],
    arsip = [],
  } = props;

  return (
    <div className="flex flex-col gap-5">
      <div className="grid grid-cols-5 gap-5">
        {/* Kegiatan */}
        <Section
          span="col-span-4"
          cols="grid-cols-4"
          title="Data Kegiatan"
          borderColor="yellow"
        >
          <DataItem
            count={timelineProject.length}
            label="Timeline Project"
            color="yellow"
          />
          <DataItem
            count={timelineDesktop.length}
            label="Timeline Desktop"
            color="yellow"
          />
          <DataItem
            count={booking.length}
            label="Booking Ruang Rapat"
            color="yellow"
          />
          <DataItem count={rapat.length} label="Jadwal Rapat" color="yellow" className="row-span-2"/>
          <DataItem count={cuti.length} label="Jadwal Cuti" color="yellow" />
          <DataItem count={meeting.length} label="Meeting" color="yellow" />
          <DataItem count={meetingList.length} label="Schedule" color="yellow" />
        </Section>

        {/* Rencana Kerja */}
        <Section
          span="col-span-1"
          cols="grid-cols-1"
          height="100%"
          title="Data Rencana Kerja"
          borderColor="green"
        >
          <DataItem count={project.length} label="Project" color="green"/>
        </Section>
      </div>

      <div className="grid grid-cols-4 gap-5">
        {/* Dokumen */}
        <Section
          span="col-span-2"
          cols="grid-cols-2"
          title="Data Dokumen"
          borderColor="blue"
        >
          <DataItem count={memo.length} label="Memo" color="cyan" />
          <DataItem count={beritaAcara.length} label="Berita Acara" color="cyan" />
          <DataItem count={surat.length} label="Surat" color="cyan" />
          <DataItem count={sk.length} label="SK" color="cyan" />  
          <DataItem
            count={perdin.length}
            label="Perjalanan Dinas"
            color="cyan"
            className="col-span-2"
          />
        </Section>

        {/* Data & Innformasi */}
        <Section
          span="col-span-2"
          cols="grid-cols-2"
          title="Data & Informasi"
          borderColor="red"
        >
          <DataItem count={masuk.length} label="Surat Masuk" color="red" />
          <DataItem count={keluar.length} label="Surat Keluar" color="red" />
          <DataItem count={arsip.length} label="Arsip" color="red" className="col-span-2"/>
        </Section>
      </div>
    </div>
  );
};
