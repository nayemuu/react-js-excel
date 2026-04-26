import { isValidNumber } from "../../utils/dataTypeCheck";
import { exportSalesReport } from "../../utils/ExcelGenerator/generateExcelForSalesReportCombined";

const Home = () => {
  return (
    <div>
      <div className="flex justify-center items-center mt-[100px]">
        <div className="bg-[#f2f3f8] p-8 rounded-md flex flex-col justify-center items-center">
          <button
            type="button"
            className="px-[37px] py-[12px] bg-gradient-to-b from-[#D13F96] to-[#833586] text-white rounded-[5px] text-lg font-bold leading-[21.48px] cursor-pointer"
            onClick={exportSalesReport}
          >
            Generate Excell
          </button>
        </div>
      </div>
    </div>
  );
};

export default Home;
